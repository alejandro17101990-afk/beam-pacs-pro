import json
from pathlib import Path

class DraftStorage:
    def __init__(self, path: str = "data/drafts.json"):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load_draft(self) -> str:
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                    return data.get("draft", "")
            except Exception:
                return ""
        return ""

    def save_draft(self, draft: str):
        with open(self.path, "w", encoding="utf-8") as fh:
            json.dump({"draft": draft}, fh, ensure_ascii=False, indent=2)
