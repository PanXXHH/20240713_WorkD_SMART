from mysupport.Pather.Pather3 import Pather
import os
import re


def getlastfile(path: str, type: str) -> str:
    pather = Pather(str(path))
    if not pather.exists():
        raise ValueError("路径不存在")

    for root, dirs, files in os.walk(pather.str()):
        break

    files = [filename for filename in files if re.match(
        r"^\d{4}-\d{2}-\d{2}"+f"\.{type}$", filename)]

    if len(files) == 0:
        return None
    files.sort(reverse=True)
    return pather(files[0]).str()


def append_subtasks_in_taskA(_task, func_ex_mpp, function_name: str, taskA: list, level=1):
    for subtask in _task.OutlineChildren:
        new_children_task_info: dict = func_ex_mpp.__getattribute__(
            function_name)(subtask, True)
        if new_children_task_info == False:
            # 当返回值为False说明不需要，忽略掉
            continue

        new_children_task_info.setdefault("Name", subtask.Name)
        new_children_task_info.setdefault(
            "OutlineLevel", subtask.OutlineLevel)
        taskA.append(new_children_task_info)
        append_subtasks_in_taskA(subtask, func_ex_mpp,
                                 function_name, taskA, level + 1)
    return taskA


def clear_project_tasks(active_project):
    tasks = active_project.Tasks
    for i in range(tasks.Count, 0, -1):
        try:
            task = tasks.Item(i)
            # print(task)
            task.Delete()
        except:
            pass
