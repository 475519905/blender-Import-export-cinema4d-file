bl_info = {
    "name": "Import Maxon Cinema 4D File",
    "blender": (4, 2, 0),
    "category": "Import-Export",
    "author": "475519905",
    "version": (1, 2, 1),
    "warning": "必须选择路径后才能正常使用！！！",
    "doc_url": "https://www.notion.so/10baf027b60980daa686d4b90a0e9443",
    "tracker_url": "https://space.bilibili.com/34368968",
    "description": "Import Maxon Cinema 4D (.c4d) files.",
}

import bpy
import os
import subprocess
import sys
import platform
from bpy.props import StringProperty, BoolProperty
from bpy.types import Operator, AddonPreferences
from bpy_extras.io_utils import ImportHelper
import tempfile

class ImportMaxonCinema4DPreferences(AddonPreferences):
    bl_idname = __name__

    c4d_install_path: StringProperty(
        name="Cinema 4D Installation Path",
        description="Path to the Cinema 4D installation folder",
        subtype='DIR_PATH',
        default=""
    )

    def draw(self, context):
        layout = self.layout
        layout.prop(self, "c4d_install_path")

class ImportMaxonCinema4DFile(Operator):
    bl_idname = "import_scene.maxon_cinema4d"
    bl_label = "Import Maxon Cinema 4D File"
    bl_description = "Import a Maxon Cinema 4D (.c4d) file"

    filename_ext = ".c4d"
    filter_glob: StringProperty(default="*.c4d", options={'HIDDEN'})

    # 添加 filepath 属性以支持拖拽
    filepath: StringProperty(subtype='FILE_PATH', options={'SKIP_SAVE'})

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

    @classmethod
    def poll(cls, context):
        # 如果有特定条件，可以在此修改
        return True

    def execute(self, context):
        c4d_file_path = self.filepath

        if not c4d_file_path or not c4d_file_path.lower().endswith(".c4d"):
            self.report({'ERROR'}, "Invalid file path or extension.")
            return {'CANCELLED'}

        # 使用临时文件夹来存储导出的 FBX 文件
        temp_dir = tempfile.gettempdir()
        fbx_file_path = os.path.join(temp_dir, "exported_file.fbx")

        addon_prefs = context.preferences.addons[__name__].preferences
        c4d_install_path = addon_prefs.c4d_install_path.strip()

        # 检查操作系统
        os_name = platform.system()
        if os_name == "Windows":
            c4dpy_executable = "c4dpy.exe"
        elif os_name == "Darwin":  # macOS
            c4dpy_executable = "c4dpy"
        else:
            self.report({'ERROR'}, "Unsupported operating system.")
            return {'CANCELLED'}

        # 构建 c4dpy 的完整路径
        c4dpy_path = os.path.join(c4d_install_path, c4dpy_executable)

        if not os.path.isfile(c4dpy_path):
            self.report({'ERROR'}, f"Could not find c4dpy at {c4dpy_path}")
            return {'CANCELLED'}

        self.export_c4d_to_fbx(c4dpy_path, c4d_file_path, fbx_file_path)

        # 在 Blender 中导入生成的 FBX 文件
        if not os.path.isfile(fbx_file_path):
            self.report({'ERROR'}, "FBX export failed.")
            return {'CANCELLED'}

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
        for obj in bpy.context.scene.objects:
            if obj.animation_data:
                obj.animation_data_clear()

    def delete_materials(self):
        for obj in bpy.context.scene.objects:
            if obj.data and hasattr(obj.data, 'materials'):
                obj.data.materials.clear()
        for material in bpy.data.materials:
            bpy.data.materials.remove(material)

    def export_c4d_to_fbx(self, c4dpy_path, c4d_file, fbx_file):
        script_content = f"""
import c4d
import os

def main():
    # 加载文档
    doc = c4d.documents.LoadDocument(r"{c4d_file}", c4d.SCENEFILTER_OBJECTS | c4d.SCENEFILTER_MATERIALS)
    if doc is None:
        raise RuntimeError("Failed to load the document.")

    c4d.documents.SetActiveDocument(doc)

    # 设置 FBX 导出参数
    export_settings = c4d.BaseContainer()
    # 在这里设置您的导出参数

    # 导出为 FBX
    if not c4d.documents.SaveDocument(doc, r"{fbx_file}", c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, c4d.FORMAT_FBX_EXPORT):
        raise RuntimeError("Failed to export the document to FBX.")

    print("Export successful! File saved to:", r"{fbx_file}")

if __name__ == "__main__":
    main()
"""
        with tempfile.NamedTemporaryFile(delete=False, suffix=".py") as temp_script_file:
            temp_script_file.write(script_content.encode('utf-8'))
            temp_script_file_path = temp_script_file.name

        # 调用 c4dpy 运行脚本
        try:
            result = subprocess.run([c4dpy_path, temp_script_file_path], check=True, capture_output=True, text=True)
            print(result.stdout)
            if result.stderr:
                print("Error:", result.stderr)
        except subprocess.CalledProcessError as e:
            print("An error occurred while running c4dpy:")
            print(e.output)
            self.report({'ERROR'}, "Failed to export FBX using c4dpy.")
        finally:
            os.remove(temp_script_file_path)

    def invoke(self, context, event):
        if self.filepath:
            # 如果已提供文件路径，直接执行
            return self.execute(context)
        else:
            # 打开文件选择窗口
            context.window_manager.fileselect_add(self)
            return {'RUNNING_MODAL'}

class IMPORT_SCENE_FH_maxon_cinema4d(bpy.types.FileHandler):
    bl_idname = "IMPORT_SCENE_FH_maxon_cinema4d"
    bl_label = "File Handler for Maxon Cinema 4D Import"
    bl_import_operator = "import_scene.maxon_cinema4d"
    bl_file_extensions = ".c4d"

    @classmethod
    def poll_drop(cls, context):
        # 允许在 3D 视图中拖拽
        return context.area.type == 'VIEW_3D'

def menu_func_import(self, context):
    self.layout.operator(ImportMaxonCinema4DFile.bl_idname, text="Maxon Cinema 4D File (.c4d)")

def register():
    bpy.utils.register_class(ImportMaxonCinema4DPreferences)
    bpy.utils.register_class(ImportMaxonCinema4DFile)
    bpy.utils.register_class(IMPORT_SCENE_FH_maxon_cinema4d)
    bpy.types.TOPBAR_MT_file_import.append(menu_func_import)

def unregister():
    bpy.utils.unregister_class(ImportMaxonCinema4DPreferences)
    bpy.utils.unregister_class(ImportMaxonCinema4DFile)
    bpy.utils.unregister_class(IMPORT_SCENE_FH_maxon_cinema4d)
    bpy.types.TOPBAR_MT_file_import.remove(menu_func_import)

if __name__ == "__main__":
    register()
