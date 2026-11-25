import os
import json


def get_project_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def find_config(config_name="default.json"):
    project_root = get_project_root()
    config_dir = os.path.join(project_root, "config")
    config_path = os.path.join(config_dir, config_name)
    if os.path.isfile(config_path):
        return config_path
    return None


def load_config(config_name="default.json"):
    path = find_config(config_name)
    if path is None:
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def apply_config(config, section, defaults):
    """用配置文件中的值填充 defaults 中为 None 的项。CLI 参数先于此处填入 defaults，然后本函数补全未设置的项。"""
    if not config or section not in config:
        return defaults
    section_cfg = config[section]
    result = dict(defaults)
    for key in result:
        if result[key] is None and key in section_cfg:
            result[key] = section_cfg[key]
    return result