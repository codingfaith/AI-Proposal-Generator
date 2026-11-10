import os
from dotenv import load_dotenv
from openai import OpenAI
import json
from typing import Dict, Any

load_dotenv()

api_key = os.environ.get("OPENAI_API_KEY")
if not api_key:
    raise RuntimeError("OPENAI_API_KEY is not set. Add it to .env locally or to your host's environment variables.")

client = OpenAI(api_key=api_key)


REQUIRED_KEYS = [
    "study_type",
    "industry",
    "background",
    "objectives",
    "universe",
    "methodology",
    "quality_controls",
    "timeline",
    "deliverables",
    "costing_assumptions",
    "recommended_approach",
    "risks",
    "assumptions"
]


def generate_proposal_json(rfq_text: str, client_name: str = "Client") -> Dict[str, Any]:
    system_prompt = """
You are a Senior Research Director at Plus94 Research.

You write professional, client-ready market research proposals for:
- Government departments
- Municipalities
- Banks
- Regulators
- State-owned enterprises
- Private sector clients
- International organisations

Write in a polished consulting style.

Do not mention AI.
Do not invent specific facts that are not in the RFQ.
Where information is missing, make reasonable proposal assumptions and clearly state them.
Return valid JSON only.
Do not include markdown, explanations, or text outside the JSON.
"""

    user_prompt = f"""
Client name: {client_name}

RFQ / client request:
{rfq_text}

Create a structured draft proposal.

Return a compact valid JSON object. Keep each section concise enough to fit inside the response limit.

{{
  "study_type": "string",
  "industry": "string",
  "background": "string",
  "objectives": ["string", "string", "string"],
  "universe": "string",
  "methodology": "string",
  "quality_controls": "string",
  "timeline": "string",
  "deliverables": ["string", "string", "string"],
  "recommended_approach": "string",
  "risks": ["string", "string", "string"],
  "assumptions": ["string", "string", "string"]
}}

Content guidance:
- Background: 150-250 words.
- Objectives: 5-8 detailed bullet points.
- Universe: describe target population, geography, and stakeholder groups.
- Methodology: 250-400 words with a clear recommended research design.
- Quality controls: include sampling checks, questionnaire review, fieldwork monitoring, validation, and data cleaning.
- Timeline: provide a practical week-by-week plan.
- Deliverables: provide clear project deliverables as bullet points.
- Costing assumptions: explain what pricing would depend on and recommend a pricing strategy.
- Recommended approach: explain why this approach is suitable.
- Risks: list possible project risks.
- Assumptions: list key assumptions used in preparing the draft.

Return ONLY valid JSON.
"""

    try:
        response = client.chat.completions.create(
            model="gpt-6.1-sol",
            messages=[
                {"role": "system", "content": system_prompt.strip()},
                {"role": "user", "content": user_prompt.strip()}
            ],
            max_completion_tokens=12000,
            response_format={"type": "json_object"}
        )

        content = response.choices[0].message.content.strip()
        try:
            proposal = json.loads(content)
        except json.JSONDecodeError:
            print("Invalid JSON returned:")
            print(content)
            raise RuntimeError("AI returned invalid JSON. Please try again.")

        for key in REQUIRED_KEYS:
            if key not in proposal:
                proposal[key] = "" if key not in ["objectives", "deliverables", "risks", "assumptions"] else []

        return proposal

    except json.JSONDecodeError as e:
        print(f"JSON Decode Error: {e}")
        print(f"Raw response was: {content[:500] if 'content' in locals() else 'No content'}")
        raise RuntimeError("AI returned invalid JSON. Please try again.")

    except Exception as e:
        print(f"OpenAI Error: {e}")
        raise RuntimeError(f"Failed to generate proposal JSON: {str(e)}")