import bpy
import os
import subprocess
from bpy.props import StringProperty, BoolProperty
from bpy.types import Operator, AddonPreferences
from bpy_extras.io_utils import ImportHelper
import tempfile


bl_info = {
    "name": "Import Maxon Cinema 4D File",
    "blender": (4, 2, 0),
    "category": "Import-Export",
    "author": "475519905",
    "version": (1, 0, 0),
    "description": "Import Cinema 4D (.c4d) files and export to c4d, with options to include/exclude models, lights, cameras, splines, animations, and materials.",
}

class ImportMaxonCinema4DPreferences(AddonPreferences):
    bl_idname = __name__

    c4d_install_path: StringProperty(
        name="Cinema 4D Installation Path",
        description="Path to the Cinema 4D installation folder",
        subtype='DIR_PATH',
        default="C:\\Program Files\\Maxon Cinema 4D 2023"
    )

    def draw(self, context):
        layout = self.layout
        layout.prop(self, "c4d_install_path")

class ImportMaxonCinema4DFile(Operator, ImportHelper):
    bl_idname = "import_scene.maxon_cinema4d"
    bl_label = "Import Maxon Cinema 4D File"

    filename_ext = ".c4d"
    filter_glob: StringProperty(default="*.c4d", options={'HIDDEN'})

    import_models: BoolProperty(name="Import Models", default=True)
    import_lights: BoolProperty(name="Import Lights", default=True)
    import_cameras: BoolProperty(name="Import Cameras", default=True)
    import_splines: BoolProperty(name="Import Splines", default=True)
    import_animations: BoolProperty(name="Import Animations", default=True)
    import_materials: BoolProperty(name="Import Materials", default=True)

    def draw(self, context):
        layout = self.layout
        layout.prop(self, "import_models")
        layout.prop(self, "import_lights")
        layout.prop(self, "import_cameras")
        layout.prop(self, "import_splines")
        layout.prop(self, "import_animations")
        layout.prop(self, "import_materials")

    def execute(self, context):
        c4d_file_path = self.filepath
        documents_path = os.path.expanduser("~/Documents")
        fbx_file_path = os.path.join(documents_path, "exported_file.fbx")

        addon_prefs = context.preferences.addons[__name__].preferences
        c4d_install_path = addon_prefs.c4d_install_path
        c4dpy_path = os.path.join(c4d_install_path, "c4dpy.exe")

        self.export_c4d_to_fbx(c4dpy_path, c4d_file_path, fbx_file_path)

        # 在 Blender 中导入生成的 FBX 文件
        bpy.ops.import_scene.fbx(filepath=fbx_file_path, use_image_search=True, use_custom_normals=True)

        # 根据用户选择删除对应类型的资产
        if not self.import_models:
            self.delete_objects_of_type('MESH')
        if not self.import_lights:
            self.delete_objects_of_type('LIGHT')
        if not self.import_cameras:
            self.delete_objects_of_type('CAMERA')
        if not self.import_splines:
            self.delete_objects_of_type('CURVE')
        if not self.import_animations:
            self.delete_animations()
        if not self.import_materials:
            self.delete_materials()

        self.report({'INFO'}, "Export and Import successful!")
        return {'FINISHED'}

    def delete_objects_of_type(self, obj_type):
        bpy.ops.object.select_all(action='DESELECT')
        for obj in bpy.context.scene.objects:
            if obj.type == obj_type:
                obj.select_set(True)
        bpy.ops.object.delete()

    def delete_animations(self):
        bpy.ops.object.select_all(action='DESELECT')
        for obj in bpy.context.scene.objects:
            if obj.animation_data:
                obj.animation_data_clear()

    def delete_materials(self):
        bpy.ops.object.select_all(action='DESELECT')
        for obj in bpy.context.scene.objects:
            obj.select_set(True)
            if obj.data and hasattr(obj.data, 'materials'):
                obj.data.materials.clear()
        bpy.ops.object.select_all(action='DESELECT')
        for material in bpy.data.materials:
            bpy.data.materials.remove(material)

    def export_c4d_to_fbx(self, c4dpy_path, c4d_file, fbx_file):
        script_content = f"""
import c4d
import os

def main():
    # 获取当前文档
    doc = c4d.documents.LoadDocument(r"{c4d_file}", c4d.SCENEFILTER_OBJECTS | c4d.SCENEFILTER_MATERIALS)
    if doc is None:
        raise RuntimeError("No document is currently open.")

    # 定义 FBX 文件导出的路径
    documents_path = os.path.expanduser("~/Documents")
    fbx_file_path = os.path.join(documents_path, "exported_file.fbx")

    # 设置 FBX 导出参数
    export_settings = c4d.BaseContainer()
    export_settings.SetInt32(c4d.FBXEXPORT_FBX_VERSION, c4d.FBX_EXPORTVERSION_NATIVE)
    export_settings.SetBool(c4d.FBXEXPORT_ASCII, False)
    export_settings.SetBool(c4d.FBXEXPORT_CAMERAS, True)
    export_settings.SetBool(c4d.FBXEXPORT_LIGHTS, True)
    export_settings.SetBool(c4d.FBXEXPORT_SPLINES, True)
    export_settings.SetBool(c4d.FBXEXPORT_INSTANCES, True)
    export_settings.SetBool(c4d.FBXEXPORT_SELECTION_ONLY, False)
    export_settings.SetBool(c4d.FBXEXPORT_GLOBAL_MATRIX, True)
    export_settings.SetInt32(c4d.FBXEXPORT_SDS, c4d.FBXEXPORT_SDS_GENERATOR)
    export_settings.SetBool(c4d.FBXEXPORT_TRIANGULATE, True)
    export_settings.SetBool(c4d.FBXEXPORT_SAVE_NORMALS, True)
    export_settings.SetBool(c4d.FBXEXPORT_SAVE_VERTEX_COLORS, True)
    export_settings.SetBool(c4d.FBXEXPORT_SAVE_VERTEX_MAPS_AS_COLORS, False)
    export_settings.SetInt32(c4d.FBXEXPORT_UP_AXIS, c4d.FBXEXPORT_UP_AXIS_Y)
    export_settings.SetBool(c4d.FBXEXPORT_FLIP_Z_AXIS, False)
    export_settings.SetBool(c4d.FBXEXPORT_TRACKS, True)
    export_settings.SetBool(c4d.FBXEXPORT_BAKE_ALL_FRAMES, True)
    export_settings.SetBool(c4d.FBXEXPORT_PLA_TO_VERTEXCACHE, True)
    export_settings.SetBool(c4d.FBXEXPORT_BOUND_JOINTS_ONLY, True)
    export_settings.SetInt32(c4d.FBXEXPORT_TAKE_MODE, c4d.FBXEXPORT_TAKE_NONE)
    export_settings.SetBool(c4d.FBXEXPORT_MATERIALS, True)
    export_settings.SetBool(c4d.FBXEXPORT_EMBED_TEXTURES, True)
    export_settings.SetBool(c4d.FBXEXPORT_SUBSTANCES, True)
    export_settings.SetBool(c4d.FBXEXPORT_BAKE_MATERIALS, True)
    export_settings.SetInt32(c4d.FBXEXPORT_BAKEDTEXTURE_WIDTH, 1024)
    export_settings.SetInt32(c4d.FBXEXPORT_BAKEDTEXTURE_HEIGHT, 1024)
    export_settings.SetInt32(c4d.FBXEXPORT_BAKEDTEXTURE_DEPTH, c4d.FBXEXPORT_BAKEDTEXTURE_DEPTH_16)
    export_settings.SetBool(c4d.FBXEXPORT_LOD_SUFFIX, False)

    # 导出 FBX 文件
    if not c4d.documents.SaveDocument(doc, fbx_file_path, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, c4d.FORMAT_FBX_EXPORT):
        raise RuntimeError("Failed to export the document to FBX.")

    print("Export successful! File saved to:", fbx_file_path)

if __name__ == "__main__":
    main()
"""

        with tempfile.NamedTemporaryFile(delete=False, suffix=".py") as temp_script_file:
            temp_script_file.write(script_content.encode('utf-8'))
            temp_script_file_path = temp_script_file.name

        subprocess.run([c4dpy_path, temp_script_file_path], check=True)
        os.remove(temp_script_file_path)

def menu_func_import(self, context):
    self.layout.operator(ImportMaxonCinema4DFile.bl_idname, text="Maxon Cinema 4d File(.c4d)")

def register():
    bpy.utils.register_class(ImportMaxonCinema4DPreferences)
    bpy.utils.register_class(ImportMaxonCinema4DFile)
    bpy.types.TOPBAR_MT_file_import.append(menu_func_import)

def unregister():
    bpy.utils.unregister_class(ImportMaxonCinema4DPreferences)
    bpy.utils.unregister_class(ImportMaxonCinema4DFile)
    bpy.types.TOPBAR_MT_file_import.remove(menu_func_import)

if __name__ == "__main__":
    register()
