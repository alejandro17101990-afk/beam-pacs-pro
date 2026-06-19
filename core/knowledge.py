import json
from pathlib import Path

class KnowledgeBase:
    def __init__(self, path: str):
        self.path = Path(path)
        self.data = self._load()

    def _load(self) -> dict:
        if self.path.exists():
            with open(self.path, "r", encoding="utf-8") as fh:
                return json.load(fh)
        return {
            "templates": {},
            "classifications": {},
            "regions": {},
            "recommendations": {},
            "definitions": {},
            "slash_commands": {},
            "differentials": {},
        }

    def list_templates(self) -> list[str]:
        return sorted(self.data.get("templates", {}).keys())

    def render_template(self, name: str) -> str:
        template = self.data.get("templates", {}).get(name, "")
        if isinstance(template, str):
            return template
        return "\n".join(template)

    def prompt_context(self) -> str:
        lines = ["Base de conocimiento radiológica activa:"]
        for key in ["classifications", "regions", "recommendations"]:
            item = self.data.get(key, {})
            if isinstance(item, dict):
                lines.append(f"- {key}: {', '.join(item.keys())}")
            else:
                lines.append(f"- {key}: {item}")
        return "\n".join(lines)

    def summary_for_context(self, report_text: str) -> str:
        region = self.detect_region(report_text)
        if region:
            return f"Contexto activo para {region}: recomendaciones de clasificación, omisiones y definiciones relevantes."
        return "Asistente contextual: ofrece clasificación, definiciones y recomendaciones inmediatas. Usa comandos slash, inserta plantillas y revisa omisiones anatómicas."

    def get_slash_commands(self) -> dict:
        return self.data.get("slash_commands", {})

    def detect_region(self, report_text: str) -> str | None:
        for region in self.data.get("regions", {}).keys():
            if region.lower() in report_text.lower():
                return region
        return None

    def get_recommendation_snippet(self, report_text: str) -> str:
        region = self.detect_region(report_text)
        if not region:
            return "No se detectó región clara; especifica indicación o región anatómica para recomendaciones precisas."
        region_info = self.data.get("regions", {}).get(region, {})
        return region_info.get("recommendation_note", "Revisar clasificaciones específicas y seguimiento según hallazgos.")

    def get_definition(self, term: str) -> str:
        return self.data.get("definitions", {}).get(term, "Definición no disponible en la base de conocimiento.")

    def get_teacher_notes(self, report_text: str) -> list[str]:
        notes = []
        for term, definition in self.data.get("definitions", {}).items():
            if term.lower() in report_text.lower():
                notes.append(f"{term}: {definition}")
        if "stoller" in report_text.lower() and "menisco" in report_text.lower():
            notes.append("Recuerda especificar si el desgarro meniscal es horizontal, vertical o complejo y su localización precisa.")
        if "tirads" in report_text.lower() or "birads" in report_text.lower():
            notes.append("Incluye siempre la categoría final y la recomendación de manejo basada en la clasificación.")
        return notes

    def get_differentials(self, report_text: str) -> list[str]:
        findings = []
        for key, values in self.data.get("differentials", {}).items():
            if key.lower() in report_text.lower():
                findings.extend(values)
        return findings

    def check_omissions(self, report_text: str) -> list[str]:
        alerts = []
        detected_region = None
        for candidate in self.data.get("regions", {}).keys():
            if candidate.lower() in report_text.lower():
                detected_region = candidate
                break
        if not detected_region:
            return alerts
        region_info = self.data.get("regions", {}).get(detected_region, {})
        for key in region_info.get("mandatory_sections", []):
            if key.lower() not in report_text.lower():
                alerts.append(f"Posible omisión en {detected_region}: revisar {key}.")
        return alerts
