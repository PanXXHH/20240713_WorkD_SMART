import os
from re import Match
from mysupport.Pather import Pather

#         print(f"任务名: {task.Name}")
#         print(f"开始时间: {task.Start}")
#         print(f"结束时间: {task.Finish}")
#         print(f"持续时间: {task.Duration} 小时")  # 假设持续时间单位为“小时”
#         print(f"完成进度: {task.PercentComplete}%")  # 输出任务的完成进度
#         print(f"备注: {task.Notes}")  # 输出任务的备注
#         print(f"大纲数字: {task.OutlineNumber}")  # 输

"""
规范手册

返回值为False时，不保留
"""

class FUNC_EX_MPP:
    def __init__(self, path: Pather) -> None:
        self.path = path
        self.selfpath = Pather(os.path.dirname(__file__))

    def _getlastfile(self):
        for root, dirs, files in os.walk(str(self.path)):
            break

        files = [filename for filename in files if re.match(
            r"^\d{4}-\d{2}-\d{2}.mpp$", filename)]
        if len(files) == 0:
            return False
        files.sort(reverse=True)
        return self.path(files[0])

    def _get_methods(self):
        return (list(filter(lambda m: not m.startswith("_") and callable(getattr(self, m)),
                            dir(self))))

    # 默认保留Name，OutlineLevel
    def SYS_DEFAULT(self, task, ischildren: bool = False, **kwargs) -> dict:
        return {"Notes": task.Notes,
                "Start": task.Start,
                "PercentComplete": task.PercentComplete}
    # 默认保留Name，OutlineLevel

    def PIN(self, task, ischildren: bool = False, **kwargs) -> dict:
        # 已完成的就去掉
        if task.PercentComplete == 100 and not ischildren:
            return False

        return {"Notes": task.Notes,
                "Start": task.Start,
                "PercentComplete": task.PercentComplete}
    # 默认保留Name，OutlineLevel

    def DEFINE(self, task, ischildren: bool = False, **kwargs) -> dict:
        return {"Notes": task.Notes}

    # 默认保留Name，OutlineLevel
    def DEFINE_FORCE(self, task, ischildren: bool = False, **kwargs) -> dict:
        if task.OutlineLevel > 1:
            task_top_parent = task.OutlineParent
            while task_top_parent.OutlineLevel != 1:
                task_top_parent = task_top_parent.OutlineParent
        else:
            task_top_parent = task

        if task_top_parent.PercentComplete == 100:
            return {"Notes": task.Notes}

        return {"Notes": task.Notes,
                "Start": task.Start,
                "PercentComplete": task.PercentComplete}

    # 默认保留Name，OutlineLevel
    def DEFINE_FORCE(self, task, ischildren: bool = False, **kwargs) -> dict:
        if task.OutlineLevel > 1:
            task_top_parent = task.OutlineParent
            while task_top_parent.OutlineLevel != 1:
                task_top_parent = task_top_parent.OutlineParent
        else:
            task_top_parent = task

        if task_top_parent.PercentComplete == 100:
            return {"Notes": task.Notes}

        return {"Notes": task.Notes,
                "Start": task.Start,
                "PercentComplete": task.PercentComplete}

    # 默认保留Name，OutlineLevel
    def DEFINE_INDEX(self, task, ischildren: bool = False, **kwargs) -> dict:
        match: Match = kwargs["match"]
        if not match:
            return False

        index = int(match.group(2))
        if index < 1:
            return False

        callback = f"{match.group(1)}({int(match.group(2))-1})"
        return {"Notes": callback}

# a = FUNC_EX_MPP(Pather(r"G:\Program Data\ONEDRIVE\MY\OneDrive - Radomil Deanne\PY-MYEXECUTION\工作-D级"))
# print(a.DEFINE_INDEX(), "11")
