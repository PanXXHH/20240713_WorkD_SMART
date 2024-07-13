from mysupport.Pather.Pather3 import Pather
from mysupport.MSProjectEx import MSProjectEx

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

class WorkDModuleUnfinishedTasks:
    def __init__(self, path) -> None:
        """
        初始化方法。

        参数:
        - path: 文件路径。
        """
        self.__path = path
        self.__important_taskinfo_keys = ["Name", "OutlineLevel"]
        self.__main_taskinfo_keys = ["Notes", "Start", "PercentComplete"]
        self.__child_taskinfo_keys = list(self.__main_taskinfo_keys)
        self.task_info_list = list()

    def old_task_filter_event(self, task):
        """
        旧的任务筛选方法。

        参数:
        - task: 需要筛选的任务。
        """
        if task.OutlineLevel > 1:
            return

        if task.Notes == "":
            if task.PercentComplete != 100:
                self.task_info_list += task_to_complete_task_info_list_by_rules(
                    task, self.__important_taskinfo_keys+self.__main_taskinfo_keys, self.__important_taskinfo_keys+self.__child_taskinfo_keys)

        # print(self.task_info_list)

    def not_func_event(self, task):
        """
        筛选未完成的任务。

        参数:
        - task: 需要筛选的任务。
        """
        if task.PercentComplete != 100:
            self.task_info_list += task_to_complete_task_info_list_by_rules(
                task, self.__important_taskinfo_keys+self.__main_taskinfo_keys, self.__important_taskinfo_keys+self.__child_taskinfo_keys)

        # print(self.task_info_list)

    def save(self, project_com: MSProjectEx, filename: str):
        """
        将筛选后的任务信息保存到新的项目文件中。

        参数:
        - project_com (MSProjectEx): MS Project的实例。
        - filename (str): 保存的文件名。
        """
        unfinishedtasks_project_file = project_com.FileOpen(str(Pather(self.__path)("UnfinishedTasks.mpp")), True)

        for task_info in self.task_info_list:
            task = unfinishedtasks_project_file.Tasks.Add(Name=task_info['Name'])
            for key, value in task_info.items():
                if key == "Name":
                    continue
                setattr(task, key, value)

        unfinishedtasks_project_file.SaveAs()
        project_com.FileClose(unfinishedtasks_project_file)
