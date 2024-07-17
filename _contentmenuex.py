import _contentmenuex_config as _config
from mysupport.Pather.Pather3 import Pather
import datetime
import os
import shutil
import re
from mysupport.MSProjectEx import MSProjectEx
import FUNC_EX_MPP
import subprocess
import libs.WorkDFSupport as utils
import libs.WorkDCProgressBarUI as ProgressBarUI
from modules.WorkDModuleUnfinishedTasks import WorkDModuleUnfinishedTasks
from modules.WorkDModulePeriodicTasks import WorkDModulePeriodicTasks
import tempfile

menu = list[tuple]()

# test # test #test

@_config.添加工作文件
def 添加工作文件(path: str, config: dict = {}):
    # 初始化进度条UI
    pb = ProgressBarUI.WorkDCProgressBarUI()
    pb.popup("Work-d：即将生成工作文件...")

    # 根据当前日期生成.mpp文件名
    now = datetime.datetime.now()
    filename = "%s.mpp" % now.strftime(r"%Y-%m-%d")

    # 如果对应日期的.mpp文件已存在，直接打开并返回
    if Pather(path)(filename).exists():
        pb.update_progress(100, "文件已存在！正在打开")
        subprocess.run(["start", "", Pather(path)(filename).str()], shell=True)
        return
    # 若.mpp文件不存在，则复制最近的文件中的任务到新文件中
    else:

        # 创建MS Project应用程序对象
        project = MSProjectEx()
        project.Visible = False

        # 获取最新的.mpp文件
        lastfile = utils.getlastfile(Pather(path), "mpp")
        if lastfile == None:
            # 若不存在最新文件则新建.mpp文件
            project.Projects.Add().SaveAs(Pather(path)(filename).str())
            project.Visible = True
            return
        else:

            # 若存在最新文件，打开并进行任务的提取与处理
            old_project_file = project.FileOpen(lastfile)
            pb.update_progress(10, "获取任务配置信息并创建新任务...")

            _func_ex_mpp = FUNC_EX_MPP.FUNC_EX_MPP(Pather(path).str())
            wdmut_1 = WorkDModuleUnfinishedTasks(Pather(path).str())
            wdmpt_2 = WorkDModulePeriodicTasks(Pather(path).str())

            # 初始化存储任务的列表
            new_tasksA = list()
            obj_new_taskAT = list()
            tasks = old_project_file.Tasks
            match = None

            # 筛选事件
            def filter_rules(task) -> bool:

                if task.OutlineLevel > 1:  # 跳过非1级任务
                    return False
                elif task.Notes == "":  # 跳过空备注任务
                    return False

                return True

            # 逐一检查任务
            for task in tasks:
                # 硬性筛选
                if task is None:
                    continue

                # 自定义事件
                wdmut_1.old_task_filter_event(task)

                # 自定义筛选规则
                if not filter_rules(task):
                    continue

                # 匹配任务备注与配置中的主机
                func = None
                for key, value in config.get("hosts", {}).items():
                    match = re.search(key, task.Notes)
                    if match:
                        func = value
                        break

                # 若未匹配到任何方法则采用默认方法
                if not func:
                    func = config.get("sys_hosts", {}).get("DEFAULT", None)
                    # 如果默认方法未定义或不在支持的方法列表中则忽略
                    if not func or func not in _func_ex_mpp._get_methods():
                        continue

                    # 以下是处理未完成任务
                    wdmut_1.not_func_event(task)

                    continue

                # 检查方法是否在支持的方法列表中
                if func not in _func_ex_mpp._get_methods():
                    raise ValueError("方法未定义")

                obj_new_taskAT.append((task, func, match))

            for task, func, match in obj_new_taskAT:
                task_info: dict = _func_ex_mpp.__getattribute__(
                    func)(task, False, match=match)

                if task_info == False or task_info == None:
                    # 当返回值为False说明不需要该值，忽略掉
                    continue

                task_info.setdefault("Name", task.Name)
                task_info.setdefault(
                    "OutlineLevel", task.OutlineLevel)
                new_tasksA.append(task_info)

                utils.append_subtasks_in_taskA(
                    task, _func_ex_mpp, func, new_tasksA)

            # 关闭旧文件，打开周期性事件文件
            project.FileClose(old_project_file)
            wdmpt_result = wdmpt_2.save(project, "PeriodicEvents.mpp")
            if wdmpt_result:
                new_tasksA += wdmpt_result

            wdmut_1.save(project, "UnfinishedTasks.mpp")

            # 代码结束 ##

            # 创建一个临时文件，并获取其路径
            with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                temp_path = temp_file.name
                print("临时文件", temp_path)
            # shutil.copy(lastfile, str(Pather(path)(filename)))
            shutil.copy(lastfile, temp_path)

            new_project_file = project.FileOpen(temp_path)

            pb.update_progress(30, "清空旧任务...")

            utils.clear_project_tasks(new_project_file)
            new_project_file.ProjectStart = datetime.datetime.now() + \
                datetime.timedelta(hours=8)

            # 将新任务添加到新文件中
            for index, task_info in enumerate(new_tasksA):
                # 更新进度条
                progress = int((index / len(new_tasksA)) * 100) * 0.7 + 30

                task_name = task_info['Name']
                task_info_text = f"正在处理任务: {task_name}"

                pb.update_progress(progress, task_info_text)

                task = new_project_file.Tasks.Add(Name=task_name)

                for key, value in task_info.items():
                    if key == "Name":
                        continue
                    setattr(task, key, value)

            pb.update_progress(100, "任务完成！")

    # 保存新文件并显示MS Project应用程序
    new_project_file.SaveAs()
    project.FileClose(new_project_file)
    project.Visible = True
    shutil.move(temp_path, str(Pather(path)(filename)))
    # project.FileOpen(str(Pather(path)(filename)))
    subprocess.run(["start", "", Pather(path)(filename).str()], shell=True)


@_config.folder_limit
def 整理文件夹(path: str, config: dict = {}):
    import PySimpleGUI as sg

    for root, dirs, files in os.walk(path):
        break

    _files = list()
    for match in config.get("matchlist"):
        _files = _files + [file for file in files if re.match(
            match, file) != None]

    # 排序
    if len(_files) <= config.get("limit"):
        sg.popup("当前文件夹不需要整理", font="微软雅黑")
        return
    _files.sort()
    # 取出要整理的文件
    passfiles = _files[:-config.get("limit")]
    if not os.path.exists(os.path.join(path, "%s\\" % config.get("dirname"))):
        os.mkdir(os.path.join(path, config.get("dirname")))
    for file in passfiles:
        shutil.move(os.path.join(path, file),
                    os.path.join(path, "%s\\" % config.get("dirname"), file))
    sg.popup("成功移动%s个文件至%s" %
             (len(passfiles), config.get("dirname")), font="微软雅黑")


# 用于外部调用
def order_release(path: str, config: dict = {}):
    # 根据当前日期生成.mpp文件名
    now = datetime.datetime.now()
    filename = "%s.mpp" % now.strftime(r"%Y-%m-%d")

    # 如果对应日期的.mpp文件已存在，直接打开并返回
    if Pather(path)(filename).exists():
        subprocess.run(["start", "", Pather(path)(filename).str()], shell=True)
        return

    添加工作文件(path=path)
    整理文件夹(path=path)
    return


menu = [
    ("添加工作文件", 添加工作文件),
    ("整理文件夹", 整理文件夹),
    ("添加工作文件并整理文件夹", lambda path:(添加工作文件(path), 整理文件夹(path))),
]

# 添加工作文件(
#     r"G:\Program Data\ONEDRIVE\987384390\OneDrive\新存储\PY-MYEXECUTION\工作-D级\.sandbox")
