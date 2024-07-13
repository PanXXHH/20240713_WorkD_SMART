import threading
import queue
import PySimpleGUI as sg


class WorkDCProgressBarUI(threading.Thread):
    def __init__(self) -> None:
        super().__init__(daemon=True)
        self.queue = queue.Queue()
        self.cancelled = False

    def _calculate_visual_length(self, text: str) -> int:
        return sum(1.7 if '\u4e00' <= char <= '\u9fff' else 1 for char in text)

    def run(self):
        text_size = 30
        layout = [
            [sg.Text('进度', justification='c', font=("Microsoft YaHei", 10))],
            [sg.ProgressBar(100, orientation='h', size=(20, 10), key='progressbar')],
            [sg.Text(size=(text_size, 1), key='-TASK_INFO-', justification='c', font=("Microsoft YaHei", 10))],
            [sg.Cancel(font=("Microsoft YaHei", 10))]
        ]
        window = sg.Window('任务进度', layout, element_justification="center", font=("Microsoft YaHei", 10))
        progress_bar = window['progressbar']

        while True:
            event, values = window.read(timeout=10)
            if event == 'Cancel' or event == sg.WINDOW_CLOSED:
                self.cancelled = True
                break

            try:
                progress, description = self.queue.get_nowait()
                text_length = self._calculate_visual_length(description)
                if text_length > text_size:
                    text_size = int(text_length)
                    window['-TASK_INFO-'].set_size((text_size, 1))

                progress_bar.update(progress)
                window['-TASK_INFO-'].update(description)
            except queue.Empty:
                pass

        window.close()

    def update_progress(self, value: int, description: str = None):
        if self.cancelled:
            raise Exception("用户取消了操作")

        if value < 0 or value > 100:
            raise ValueError("给定的value超出范围")
        self.queue.put((value, description))

    def popup(self, description: str = None):
        self.start()
        self.update_progress(0, description)


# if __name__ == '__main__':
#     progress_ui = WorkDCProgressBarUI()
#     progress_ui.popup("开始任务")
