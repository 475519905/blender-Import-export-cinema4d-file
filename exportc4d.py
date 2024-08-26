import bpy
import os
import subprocess
import tempfile
import logging
from bpy_extras.io_utils import ExportHelper
from bpy.props import StringProperty, BoolProperty
from bpy.types import Operator, AddonPreferences

bl_info = {
    "name": "Export Maxon Cinema 4D File",
    "author": "475519905",
    "version": (1, 8),
    "blender": (4, 2, 0),
    "location": "File > Export > Cinema 4D (.c4d)",
    "description": "Export the scene to Cinema 4D format via FBX",
    "category": "Import-Export",
}

# 设置日志
log_file = os.path.join(os.path.expanduser("~"), "Documents", "blender_c4d_export_log.txt")
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    filename=log_file,
                    filemode='a')
logger = logging.getLogger(__name__)

class exportMaxonCinema4DPreferences(AddonPreferences):
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

class ExportC4D(Operator, ExportHelper):
    bl_idname = "export_scene.c4d"
    bl_label = "Export C4D"
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
            with open(c4d_script_path, 'w') as f:
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

def menu_func_export(self, context):
    self.layout.operator(ExportC4D.bl_idname, text="Maxon Cinema 4d File(.c4d)")

def register():
    bpy.utils.register_class(ExportC4D)
    bpy.utils.register_class(exportMaxonCinema4DPreferences)
    bpy.types.TOPBAR_MT_file_export.append(menu_func_export)

def unregister():
    bpy.utils.unregister_class(ExportC4D)
    bpy.utils.unregister_class(exportMaxonCinema4DPreferences)
    bpy.types.TOPBAR_MT_file_export.remove(menu_func_export)

if __name__ == "__main__":
    register()
