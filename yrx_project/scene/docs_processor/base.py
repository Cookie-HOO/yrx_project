from multiprocessing import Lock


class ActionContext:
    doc = None
    selection = None
    word = None
    file_path = None

    selected_range = None  # 经过select类型的cmd，就会被赋值

    msg = ""
    done = 0
    total_task = -1

    def __init__(self):
        self.doc = None
        self.selection = None
        self.word = None
        self.file_path = None
        self.msg = None
        self.done = 0
        self.total_task = -1

        self.lock = Lock()

    def done_task(self):  # 多进程级别保证同步
        with self.lock:
            self.done += 1


class Command:
    action_type = None
    action_name = None

    def __init__(self, action_type, action_name, **kwargs):
        self.action_type = action_type
        self.action_name = action_name

    def run(self, context: ActionContext):
        return self.office_word_run(context)

    def office_word_run(self, context: ActionContext):
        raise NotImplementedError

