from mysupport.Pather.Pather3 import Pather
from mysupport.MSProjectEx import MSProjectEx
import datetime
import subprocess


def task_to_task_info_by_rules(task: dict, rules: list) -> dict:
    """
    根据提供的规则从任务中提取信息。

    参数:
    - task (dict): 需要提取信息的任务。
    - rules (list): 需要从任务中提取的属性列表。

    返回:
    - dict: 提取的任务信息。
    """
    if not rules:
        return None

    task_info = dict()
    for key in rules:
        task_info[key] = getattr(task, key)

    return task_info


def get_subtask_info_list_by_task(task, rules: list) -> list:
    """
    获取任务的子任务信息列表。

    参数:
    - task: 主任务。
    - rules (list): 需要从子任务中提取的属性列表。

    返回:
    - list: 子任务信息列表。
    """
    subtask_info_list = list()

    def __process_subtasks(task, rules: list, results: list, level: int = 1):
        for subtask in task.OutlineChildren:
            subtask_info: dict = task_to_task_info_by_rules(subtask, rules)
            if subtask_info == None:
                continue

            results.append(subtask_info)
            __process_subtasks(subtask, rules, results, level + 1)

    __process_subtasks(task, rules, subtask_info_list)

    return subtask_info_list


def task_to_complete_task_info_list_by_rules(task: dict, main_task_rules: list, sub_task_rules: list) -> list:
    """
    根据规则从任务中提取完整的任务信息列表。

    参数:
    - task (dict): 主任务。
    - main_task_rules (list): 主任务的属性规则列表。
    - sub_task_rules (list): 子任务的属性规则列表。

    返回:
    - list: 完整的任务信息列表。
    """
    return [task_to_task_info_by_rules(task, main_task_rules)] + get_subtask_info_list_by_task(task, sub_task_rules)


class WorkDModulePeriodicTasks:
    def __init__(self, path) -> None:
        """
        初始化方法。

        参数:
        - path: 文件路径。
        """
        self.__path = path
        self.__important_taskinfo_keys = ["Name", "OutlineLevel"]
        self.__main_taskinfo_keys = ["Notes"]
        self.__sub_taskinfo_keys = list(self.__main_taskinfo_keys)
        self.task_info_list = list()

    def save(self, project_com: MSProjectEx, filename: str):
        """
        将筛选后的任务信息保存到新的项目文件中。

        参数:
        - project_com (MSProjectEx): MS Project的实例。
        - filename (str): 保存的文件名。
        """
        periodicevents_project_file = project_com.FileOpen(
            str(Pather(self.__path)("PeriodicEvents.mpp")), True)

        # 获取今天的日期
        today = datetime.datetime.today().date()

        # 更新周期性事件文件中的任务开始日期
        for task in periodicevents_project_file.Tasks:
            if task == None:
                continue
            
            if task.OutlineLevel > 1:
                continue

            start1: datetime.datetime = task.Start1
            if start1 != "NA" and start1.date() < today:
                if task.Notes == "":
                    task.Delete()
                    continue

                start = task.Start.replace(tzinfo=None)
                start_tz = task.Start.tzinfo
                relative_days = (
                    start - datetime.datetime(1899, 12, 30)).days

                # Run the script and pass relative_days as argument
                # Replace with your actual script path
                script_path = rf"extends://daymatter?func%3Dnext_month_fixed_day%26start_date%3D{relative_days}"
                cmd = ["python", Pather(self.__path)(".scripts\\")(
                    "DayMatter\\")("__init__.py").str(), script_path]
                result = subprocess.run(cmd, stdout=subprocess.PIPE)
                # parse the result
                output = result.stdout.decode('utf-8').strip()
                task.Start1 = datetime.datetime.fromisoformat(
                    output).replace(tzinfo=start_tz)

        result = []

        # 从周期性事件文件中提取当天任务
        for task in periodicevents_project_file.Tasks:
            if task == None:
                continue
            
            if task.OutlineLevel > 1:
                continue
            
            if task.Start1 == "NA":
                continue
            
            print(task)
            start1: datetime.datetime = task.Start1
            if start1.date() == today:
                task_info = task_to_task_info_by_rules(
                    task, self.__important_taskinfo_keys+self.__main_taskinfo_keys)
                task_info["Notes"] = f"PeriodicEvents: {task.ID}"
                result.append(task_info)

                result += get_subtask_info_list_by_task(
                    task, self.__important_taskinfo_keys+self.__sub_taskinfo_keys)

            print(result)

        periodicevents_project_file.SaveAs()
        project_com.FileClose(periodicevents_project_file)

        return result
