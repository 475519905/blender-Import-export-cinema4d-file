import bpy
import os
import subprocess
import tempfile
import logging
import sys
import platform
from bpy_extras.io_utils import ExportHelper, ImportHelper
from bpy.props import StringProperty, BoolProperty
from bpy.types import Operator, AddonPreferences, FileHandler

bl_info = {
    "name": "Maxon Cinema 4D Import-Export",
    "author": "475519905",
    "version": (2, 0, 0),
    "blender": (4, 2, 0),
    "location": "File > Import-Export > Maxon Cinema 4D (.c4d)",
    "description": "Import and Export Maxon Cinema 4D (.c4d) files.",
    "category": "Import-Export",
    "support": "COMMUNITY",
    "warning": "Must select a path to use normally!",
    "doc_url": "https://space.bilibili.com/34368968",
    "tracker_url": "https://github.com/475519905/blender-Import-export-cinema4d-file/tree/main",
}


# 设置日志
log_file = os.path.join(os.path.expanduser("~"), "Documents", "blender_c4d_log.txt")
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filename=log_file,
                    filemode='a')
logger = logging.getLogger(__name__)


class MaxonCinema4DPreferences(AddonPreferences):
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

        # 第一行按钮
        row = layout.row()
        row.operator(
            "wm.url_open", 
            text="查看文档",
            icon='HELP'
        ).url = "https://www.yuque.com/shouwangxingkong-0p4w3/ldvruc/wn8pstg9pwd6r6mp?singleDoc"
        row.operator(
            "wm.url_open", 
            text="关于作者",
            icon='USER'
        ).url = "https://space.bilibili.com/34368968"

        # 第二行按钮
        row = layout.row()
        row.operator(
            "wm.url_open", 
            text="检查更新",
            icon='FILE_REFRESH'
        ).url = "https://github.com/475519905/blender-Import-export-cinema4d-file"
        row.operator(
            "wm.url_open", 
            text="加入QQ群",
            icon='COMMUNITY'
        ).url = "https://qm.qq.com/cgi-bin/qm/qr?k=9KgmVUQMfoGf7g_s-4tSe15oMJ6rbz6b&jump_from=webapi&authKey=hs9XWuCbT1jx9ytpzSsXbJuQCwUc2kXy0gRJfA+qMaVoXTbvhiOKz0dHOnP1+Cvt"

        # 第三行按钮（单个）
        row = layout.row()
        row.operator(
            "wm.url_open", 
            text="购买付费版",
            icon='FUND'
        ).url = "https://www.bilibili.com/video/BV1x2pSeNEaV/"

class ExportMaxonCinema4D(Operator, ExportHelper):
    bl_idname = "export_scene.maxon_cinema4d"
    bl_label = "Export Maxon Cinema 4D File"
    bl_options = {'PRESET', 'UNDO'}

    filename_ext = ".c4d"
    filter_glob: StringProperty(default="*.c4d", options={'HIDDEN'})

    use_selection: BoolProperty(
        name="Selected Objects",
        description="Export selected objects only",
        default=False,
    )

    def execute(self, context):
        logger.info("开始C4D导出操作")
        return self.export_c4d(context, self.filepath)

    def export_c4d(self, context, filepath):
        preferences = context.preferences.addons[__name__].preferences
        c4d_install_path = preferences.c4d_install_path

        with tempfile.TemporaryDirectory() as temp_dir:
            # 导出FBX
            fbx_path = os.path.join(temp_dir, "temp.fbx")
            logger.info(f"正在导出FBX到: {fbx_path}")
            try:
                bpy.ops.export_scene.fbx(
                    filepath=fbx_path,
                    use_selection=self.use_selection,
                    path_mode='COPY'
                )
            except Exception as e:
                logger.error(f"FBX导出失败: {str(e)}")
                self.report({'ERROR'}, f"FBX导出失败。请查看日志文件: {log_file}")
                return {'CANCELLED'}

            # 创建C4D Python脚本
            c4d_script = f"""
import c4d
import sys
import os

def main():
    print("开始执行C4D脚本")
    print(f"Cinema 4D 版本: {{c4d.GetC4DVersion()}}")
    print(f"Python 版本: {{sys.version}}")
    print(f"当前工作目录: {{os.getcwd()}}")
    
    fbx_path = r"{fbx_path}"
    if not os.path.exists(fbx_path):
        print("FBX文件不存在: " + fbx_path, file=sys.stderr)
        return

    try:
        # 创建一个新文档
        doc = c4d.documents.BaseDocument()
        
        # 导入FBX文件
        if not c4d.documents.MergeDocument(doc, fbx_path, c4d.SCENEFILTER_OBJECTS | c4d.SCENEFILTER_MATERIALS):
            print("无法导入FBX文件", file=sys.stderr)
            return
        
        # 将新文档设置为活动文档
        c4d.documents.SetActiveDocument(doc)
        
        print("FBX文件成功导入")
    except Exception as e:
        print("导入FBX时发生错误: " + str(e), file=sys.stderr)
        return

    c4d_path = r"{filepath}"
    try:
        if not c4d.documents.SaveDocument(doc, c4d_path, c4d.SAVEDOCUMENTFLAGS_DONTADDTORECENTLIST, c4d.FORMAT_C4DEXPORT):
            print("保存C4D文件失败", file=sys.stderr)
        else:
            print("C4D文件保存成功: " + c4d_path)
    except Exception as e:
        print("保存C4D文件时发生错误: " + str(e), file=sys.stderr)

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("C4D脚本错误: " + str(e), file=sys.stderr)
"""

            c4d_script_path = os.path.join(temp_dir, "import_export_c4d.py")
            with open(c4d_script_path, 'w', encoding='utf-8') as f:
                f.write(c4d_script)

            # 运行C4D Python脚本
            c4d_exe = os.path.join(c4d_install_path, "c4dpy.exe")
            logger.info(f"正在运行C4D脚本: {c4d_script_path}")
            try:
                env = os.environ.copy()
                env['MAXON_LICENSE_LOCATION'] = r'C:\ProgramData\Maxon\MAXON App\licenses'
                result = subprocess.run([c4d_exe, c4d_script_path], 
                                        check=True, 
                                        capture_output=True, 
                                        text=True,
                                        env=env)
                logger.info(f"C4D脚本输出: {result.stdout}")
                if result.stderr:
                    logger.error(f"C4D脚本错误: {result.stderr}")
            except subprocess.CalledProcessError as e:
                logger.error(f"运行C4D脚本时出错: {str(e)}")
                logger.error(f"C4D脚本输出: {e.stdout}")
                logger.error(f"C4D脚本错误: {e.stderr}")
                self.report({'ERROR'}, f"运行C4D脚本时出错。请查看日志文件: {log_file}")
                return {'CANCELLED'}

        logger.info(f"导出完成: {filepath}")
        self.report({'INFO'}, f"已导出到: {filepath}")
        self.report({'INFO'}, f"日志文件位置: {log_file}")
        return {'FINISHED'}

class ImportMaxonCinema4D(Operator, ImportHelper):
    bl_idname = "import_scene.maxon_cinema4d"
    bl_label = "Import Maxon Cinema 4D File"
    bl_description = "Import a Maxon Cinema 4D (.c4d) file"

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

    @classmethod
    def poll(cls, context):
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

        try:
            bpy.ops.import_scene.fbx(filepath=fbx_file_path, use_image_search=True, use_custom_normals=True)
        except Exception as e:
            logger.error(f"FBX导入失败: {str(e)}")
            self.report({'ERROR'}, f"FBX导入失败。请查看日志文件: {log_file}")
            return {'CANCELLED'}

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

        self.report({'INFO'}, "导入完成！")
        self.report({'INFO'}, f"日志文件位置: {log_file}")
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
        with tempfile.NamedTemporaryFile(delete=False, suffix=".py", mode='w', encoding='utf-8') as temp_script_file:
            temp_script_file.write(script_content)
            temp_script_file_path = temp_script_file.name

        # 调用 c4dpy 运行脚本
        try:
            result = subprocess.run([c4dpy_path, temp_script_file_path], check=True, capture_output=True, text=True)
            logger.info(f"c4dpy 输出: {result.stdout}")
            if result.stderr:
                logger.error(f"c4dpy 错误: {result.stderr}")
        except subprocess.CalledProcessError as e:
            logger.error(f"运行c4dpy时出错: {str(e)}")
            logger.error(f"c4dpy 输出: {e.stdout}")
            logger.error(f"c4dpy 错误: {e.stderr}")
            self.report({'ERROR'}, f"运行c4dpy时出错。请查看日志文件: {log_file}")
        finally:
            os.remove(temp_script_file_path)

    def invoke(self, context, event):
        if self.filepath:
            return self.execute(context)
        else:
            context.window_manager.fileselect_add(self)
            return {'RUNNING_MODAL'}

class ImportMaxonCinema4DFileHandler(FileHandler):
    bl_idname = "IMPORT_SCENE_FH_maxon_cinema4d"
    bl_label = "File Handler for Maxon Cinema 4D Import"
    bl_import_operator = "import_scene.maxon_cinema4d"
    bl_file_extensions = ".c4d"

    @classmethod
    def poll_drop(cls, context):
        return context.area.type == 'VIEW_3D'

def menu_func_export(self, context):
    self.layout.operator(ExportMaxonCinema4D.bl_idname, text="Maxon Cinema 4D File (.c4d)")

def menu_func_import(self, context):
    self.layout.operator(ImportMaxonCinema4D.bl_idname, text="Maxon Cinema 4D File (.c4d)")

def register():
    bpy.utils.register_class(MaxonCinema4DPreferences)
    bpy.utils.register_class(ExportMaxonCinema4D)
    bpy.utils.register_class(ImportMaxonCinema4D)
    bpy.utils.register_class(ImportMaxonCinema4DFileHandler)
    bpy.types.TOPBAR_MT_file_export.append(menu_func_export)
    bpy.types.TOPBAR_MT_file_import.append(menu_func_import)

def unregister():
    bpy.utils.unregister_class(MaxonCinema4DPreferences)
    bpy.utils.unregister_class(ExportMaxonCinema4D)
    bpy.utils.unregister_class(ImportMaxonCinema4D)
    bpy.utils.unregister_class(ImportMaxonCinema4DFileHandler)
    bpy.types.TOPBAR_MT_file_export.remove(menu_func_export)
    bpy.types.TOPBAR_MT_file_import.remove(menu_func_import)

if __name__ == "__main__":
    register()
