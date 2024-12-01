import ast


class SafeFunctionChecker(ast.NodeVisitor):
    def __init__(self):
        self.errors = []
        self.disallowed_modules = {'os'}  # 禁止导入的模块
        self.allowed_functions = {'str', 'int', 'float', 'len'}  # 允许的内置函数
        self.allowed_module_functions = {'re': {'find', 'match', 'search', 'findall'}}  # 允许的模块函数

    def visit_Import(self, node):
        for alias in node.names:
            if alias.name in self.disallowed_modules:
                self.errors.append(f"Importing module '{alias.name}' is not allowed (line {node.lineno}).")

    def visit_ImportFrom(self, node):
        if node.module in self.disallowed_modules:
            self.errors.append(f"Importing from module '{node.module}' is not allowed (line {node.lineno}).")

    def visit_Call(self, node):
        # 检查函数调用
        if isinstance(node.func, ast.Name):
            if node.func.id not in self.allowed_functions:
                self.errors.append(f"Function '{node.func.id}' is not allowed (line {node.lineno}).")
        elif isinstance(node.func, ast.Attribute):
            # 检查模块函数调用
            if isinstance(node.func.value, ast.Name):
                module_name = node.func.value.id
                function_name = node.func.attr
                if module_name in self.allowed_module_functions:
                    if function_name not in self.allowed_module_functions[module_name]:
                        self.errors.append(f"Function '{module_name}.{function_name}' is not allowed (line {node.lineno}).")
                # else:
                #     self.errors.append(f"Module '{module_name}' is not allowed (line {node.lineno}).")
        self.generic_visit(node)

    def check_code(self, code_text):
        try:
            tree = ast.parse(code_text)
            self.visit(tree)
            if self.errors:
                return False, "⚠️存在危险代码: \n" + "\n".join(self.errors)
            return True, None
        except SyntaxError as e:
            error_details = (
                f"❌语法校验失败\n"
                f"错误信息: {e.msg}\n"
                f"行号: {e.lineno}\n"
                f"偏移（列）: {e.offset}\n"
                f"错误文本: {e.text.strip() if e.text else 'N/A'}"
            )
            return False, error_details


class PythonCodeParser:
    def __init__(self, code_text, entry_func):
        self.code_text = code_text
        self.entry_func = entry_func

    def check_code(self) -> (bool, str):
        checker = SafeFunctionChecker()
        is_safe, details = checker.check_code(self.code_text)
        if not is_safe:
            return False, details
        return True, ""

    def get_func(self):
        # 定义受限的命名空间
        import re
        safe_globals = {
            "__builtins__": {k: __builtins__[k] for k in ('str', 'int', 'float', 'len', 'bool', "__import__")},
            "re": re  # 预先导入 re 模块
        }
        local_namespace = {}

        # 执行用户代码
        exec(self.code_text, safe_globals, local_namespace)

        # 获取用户定义的函数
        user_function = local_namespace.get(self.entry_func)
        if not user_function:
            raise ValueError(f"没有入口函数：{self.entry_func}")
        return user_function
