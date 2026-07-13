from typing import Any

VALUE_KEYS = ["valueString", "valueNumber", "valueDate", "valueCurrency", "content"]


def extract_value(field_payload: dict[str, Any] | None) -> Any:
    if not isinstance(field_payload, dict):
        return None
    for key in VALUE_KEYS:
        if key in field_payload and field_payload[key] not in (None, ""):
            return field_payload[key]
    return None


def extract_confidence(field_payload: dict[str, Any] | None) -> float | None:
    if not isinstance(field_payload, dict):
        return None
    confidence = field_payload.get("confidence")
    return float(confidence) if confidence is not None else None


def extract_with_fallback(raw_fields: dict[str, Any], sources: list[str]) -> dict[str, Any]:
    for source_name in sources:
        if source_name in raw_fields:
            payload = raw_fields[source_name]
            return {
                "value": extract_value(payload),
                "confidence": extract_confidence(payload),
                "source_field": source_name,
            }
    return {"value": None, "confidence": None, "source_field": None}
