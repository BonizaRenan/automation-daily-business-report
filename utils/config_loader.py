import json


def load_config(path="shared/Json/subject.json"):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)