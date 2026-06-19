import logging
from openai import OpenAI

REPORT_SYSTEM_PROMPT = """Eres un asistente de redacción de informes radiológicos. Tu tarea es transformar hallazgos clínicos en informes estructurados profesionales.

Debes generar un informe con esta estructura:
- TÉCNICA: describe el protocolo y secuencias utilizadas
- HALLAZGOS: organiza los hallazgos por estructura anatómica
- IMPRESIÓN DIAGNÓSTICA: conclusiones en viñetas

Reglas:
- Usa lenguaje preciso y terminología radiológica estándar
- Mantén la redacción clara y concisa
- No agregues información no respaldada por los hallazgos
- Si el usuario provee hallazgos incompletos, usa la estructura pero señala con "[pendiente de especificar]"
- Responde SOLO con el informe, sin explicaciones ni comentarios adicionales"""


class RadiologyCopilot:
    def __init__(self, api_key: str, knowledge_base=None):
        self.api_key = api_key
        self.knowledge_base = knowledge_base
        self.client = OpenAI(api_key=api_key, timeout=30.0) if api_key else None

    def is_available(self) -> bool:
        return self.client is not None

    def generate_report(self, text: str) -> str:
        if not self.client:
            return text
        kb_snippet = ""
        if self.knowledge_base:
            kb_snippet = self.knowledge_base.prompt_context()
        messages = [
            {"role": "system", "content": REPORT_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    f"Contexto radiológico disponible:\n{kb_snippet}\n\n"
                    f"Hallazgos del estudio:\n{text}\n\n"
                    f"Genera el informe estructurado."
                ),
            },
        ]
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=messages,
                temperature=0.15,
                max_tokens=2048,
            )
            return response.choices[0].message.content
        except Exception as e:
            logging.error(f"Error generating report: {e}")
            return f"Error al generar el informe: {e}"

    def refine_report(self, report_text: str) -> str:
        if not self.client:
            return report_text
        prompt = (
            "Revisa y mejora el siguiente informe radiológico. "
            "Corrige errores, mejora la redacción, unifica terminología y "
            "asegura que la estructura sea profesional (TÉCNICA, HALLAZGOS, IMPRESIÓN DIAGNÓSTICA). "
            "Responde SOLO con el informe mejorado, sin explicaciones.\n\n"
            f"Informe:\n{report_text}"
        )
        try:
            response = self.client.chat.completions.create(
                model="deepseek-chat",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.1,
                max_tokens=2048,
            )
            return response.choices[0].message.content
        except Exception as e:
            logging.error(f"Error refining report: {e}")
            return f"Error al refinar el informe: {e}"
