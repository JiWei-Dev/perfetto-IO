"""大模型客户端（OpenAI 兼容接口）

通过 API Key 调用大模型云端接口，为 IO 分析产物提供 AI 深度研判能力。

模型信息统一由项目配置文件 `config/ai_model.json` 管理，更换模型 / 服务商时
只需修改该配置文件，无需改动代码：
  - base_url   : 服务端点（DeepSeek 官方 / 本地 Ollama 等任意 OpenAI 兼容地址）
  - fast_model : 快速、便宜模型（结构化总结 / 指标提取）
  - deep_model : 推理模型（根因研判，输出含 reasoning_content 思考过程）

本地 Ollama 模型切换：
  - 设置环境变量 AI_MODEL_LOCAL=1 即可启用 config/ai_model.json 中的 "local" 段；
  - 或直接在 chat()/analyze_with_prompt() 传入 base_url 参数（本地端点无需 API Key）。

API Key 读取优先级: 显式传参 > 环境变量 (api_key_env 指定名称) > 项目根目录 .env 文件。
"""
import os
import json
import urllib.request
import urllib.error


def _project_root():
    # scripts/AI/llm_client.py -> 项目根目录
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _load_model_config():
    """读取 config/ai_model.json，文件缺失或损坏时返回空 dict（走内置默认值）。"""
    path = os.path.join(_project_root(), "config", "ai_model.json")
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (ValueError, OSError):
        return {}


_MODEL_CONFIG = _load_model_config()

# 本地模型预设：环境变量 AI_MODEL_LOCAL=1 时，改用 config/ai_model.json 中的 "local" 段
_LOCAL_CONFIG = _MODEL_CONFIG.get("local", {})
_USE_LOCAL = os.environ.get("AI_MODEL_LOCAL", "").strip().lower() in ("1", "true", "yes", "on")


def _pick(key, default):
    """本地预设开启时优先取 local 段，否则取顶层配置。"""
    if _USE_LOCAL and key in _LOCAL_CONFIG:
        return _LOCAL_CONFIG[key]
    return _MODEL_CONFIG.get(key, default)


# 服务端点：环境变量 DEEPSEEK_BASE_URL 优先（便于本地 Ollama 等一键切换），否则取配置文件
BASE_URL = os.environ.get(
    "DEEPSEEK_BASE_URL", _pick("base_url", "https://api.deepseek.com"))

DEFAULT_FAST_MODEL = _pick("fast_model", "deepseek-chat")
DEFAULT_DEEP_MODEL = _pick("deep_model", "deepseek-reasoner")
API_KEY_ENV = _MODEL_CONFIG.get("api_key_env", "DEEPSEEK_API_KEY")
DEFAULT_TEMPERATURE = _pick("temperature", 0.2)
DEFAULT_MAX_TOKENS = _pick("max_tokens", 8192)
DEFAULT_TIMEOUT = _pick("timeout", 600)


def _load_env_file():
    """读取项目根目录 .env 文件（简单 KEY=VALUE 解析，忽略注释与空行）。"""
    env_path = os.path.join(_project_root(), ".env")
    if not os.path.isfile(env_path):
        return {}
    env = {}
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            env[key.strip()] = val.strip().strip('"').strip("'")
    return env


def get_api_key(api_key=None):
    """按优先级返回 API Key: 显式传参 > 环境变量 > .env 文件。"""
    if api_key:
        return api_key
    env_key = os.environ.get(API_KEY_ENV, "")
    if env_key:
        return env_key
    return _load_env_file().get(API_KEY_ENV, "")


def chat(messages, model=None, role="fast", temperature=None,
         max_tokens=None, api_key=None, timeout=None, base_url=None):
    """调用 chat completions 接口，返回 (content, reasoning, usage)。

    参数:
      messages:  OpenAI 格式的 [{"role":..., "content":...}, ...]
      role:      "fast" 或 "deep"，决定默认模型与是否传 temperature
      base_url:  服务端点，默认取环境变量 / 配置文件（本地 Ollama 可传 http://localhost:11434/v1）
    未显式传入的参数均回落到 config/ai_model.json 中的配置值。
    本地端点（localhost / 127.0.0.1）无需 API Key。
    """
    url_base = (base_url or BASE_URL).rstrip("/")
    is_local = url_base.startswith("http://localhost") or url_base.startswith("http://127.0.0.1")

    key = get_api_key(api_key)
    if not key and not is_local:
        raise RuntimeError(
            "未找到 API Key（环境变量 %s），请设置环境变量或写入项目根目录 .env 文件"
            % API_KEY_ENV)

    if model is None:
        model = DEFAULT_DEEP_MODEL if role == "deep" else DEFAULT_FAST_MODEL
    if temperature is None:
        temperature = DEFAULT_TEMPERATURE
    if max_tokens is None:
        max_tokens = DEFAULT_MAX_TOKENS
    if timeout is None:
        timeout = DEFAULT_TIMEOUT

    url = url_base + "/chat/completions"
    body = {
        "model": model,
        "messages": messages,
        "stream": False,
        "max_tokens": max_tokens,
    }
    # deepseek-reasoner 不支持 temperature 等采样参数，仅 fast 模型传入
    if role != "deep":
        body["temperature"] = temperature

    headers = {"Content-Type": "application/json"}
    # 本地端点默认不需要鉴权；仅在显式传了 api_key 时才附带 Authorization
    if key and (not is_local or api_key):
        headers["Authorization"] = "Bearer " + key

    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers=headers,
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            detail = e.read().decode("utf-8")
        except Exception:
            pass
        raise RuntimeError("模型接口返回 HTTP %s: %s" % (e.code, detail))

    msg = data["choices"][0]["message"]
    content = msg.get("content", "") or ""
    reasoning = msg.get("reasoning_content", "") or ""
    usage = data.get("usage", {})
    return content, reasoning, usage


def analyze_with_prompt(system_prompt, user_prompt, role="deep",
                        max_tokens=None, temperature=None,
                        api_key=None, timeout=None, base_url=None):
    """便捷封装: 用 system + user 两段提示调用模型，返回最终 content。"""
    content, _reasoning, _usage = chat(
        [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        role=role, temperature=temperature, max_tokens=max_tokens,
        api_key=api_key, timeout=timeout, base_url=base_url,
    )
    return content