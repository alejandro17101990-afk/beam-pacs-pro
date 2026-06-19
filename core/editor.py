class RadiologyEditor:
    def __init__(self, knowledge_base=None):
        self.knowledge_base = knowledge_base

    def insert_command_template(self, report_text: str, command: str) -> str:
        if not self.knowledge_base:
            return report_text
        cmds = self.knowledge_base.get_slash_commands()
        payload = cmds.get(command, "")
        if isinstance(payload, dict):
            template = payload.get("template", "")
        else:
            template = payload
        if not template:
            return report_text
        if report_text.strip().endswith("\n") or not report_text.strip():
            return f"{report_text.strip()}\n{template}\n"
        return f"{report_text.strip()}\n\n{template}\n"
