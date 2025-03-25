import os

from yrx_project.scene.docs_processor.base import Command, ActionContext


class MergeDocumentsCommand(Command):

    def office_word_run(self, context: ActionContext):
        word = context.word
        new_doc = word.Documents.Add()

        try:
            # 创建一个范围对象，初始指向文档末尾
            range_obj = new_doc.Range(0, 0)

            for file_path in context.input_paths:
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"Merge source not found: {file_path}")

                # 打开源文件并复制内容
                doc = word.Documents.Open(os.path.abspath(file_path))
                doc.Content.Copy()
                doc.Close(SaveChanges=False)

                # 插入换行符分隔不同文件内容
                range_obj.InsertAfter("\n")  # 插入换行符或其他分隔符
                range_obj.Collapse(Direction=0)  # 移动光标到当前范围的末尾

                # 粘贴内容到指定范围
                range_obj.Paste()

                # 更新范围对象到文档末尾
                range_obj.End = new_doc.Content.End

            # 保存合并后的文档
            output_path = os.path.join(f"{context.command_container.output_folder}",
                                       f"{context.command.action_name}.docx")
            new_doc.SaveAs(os.path.abspath(output_path))
            context.input_paths = [output_path]

        finally:
            new_doc.Close()
            # 确保所有临时文档关闭
            for doc in word.Documents:
                if doc.Name != new_doc.Name:
                    doc.Close(SaveChanges=False)