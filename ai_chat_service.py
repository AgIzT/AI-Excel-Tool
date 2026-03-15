# -*- coding: utf-8 -*-
import os
from typing import Callable, List, Dict, Optional, Any


def _resolve_api_key(explicit_key: Optional[str] = None) -> str:
    if explicit_key:
        return explicit_key.strip()
    key = (os.getenv("GLM_API_KEY") or os.getenv("ZHIPU_API_KEY") or "").strip()
    if key:
        return key
    raise RuntimeError("未配置 GLM API Key，请在设置页填写后保存")


class GLMChatAssistant:
    def __init__(
        self,
        api_key: Optional[str] = None,
        base_url: Optional[str] = None,
        model: str = "glm-5",
        system_prompt: str = "你是一个专业、准确、简洁的中文AI助手。",
        max_history: int = 20,
        temperature: float = 0.3,
        max_tokens: int = 2048,
        client_factory: Optional[Callable[..., Any]] = None,
    ):
        self.api_key = _resolve_api_key(api_key)
        self.base_url = (base_url or os.getenv("GLM_BASE_URL") or "https://open.bigmodel.cn/api/paas/v4/").strip()
        self.model = (model or os.getenv("GLM_CHAT_MODEL") or "glm-5").strip()
        self.temperature = temperature
        self.max_tokens = max_tokens
        self.max_history = max(2, int(max_history))
        self.system_prompt = (system_prompt or "").strip() or "你是一个专业、准确、简洁的中文AI助手。"
        self._client_factory = client_factory or self._build_openai_client
        self.client = self._client_factory(api_key=self.api_key, base_url=self.base_url)
        self.messages: List[Dict[str, str]] = [{"role": "system", "content": self.system_prompt}]

    def reset(self, system_prompt: Optional[str] = None):
        if system_prompt is not None and system_prompt.strip():
            self.system_prompt = system_prompt.strip()
        self.messages = [{"role": "system", "content": self.system_prompt}]

    def send_message(self, user_message: str) -> str:
        user_message = (user_message or "").strip()
        if not user_message:
            raise ValueError("消息内容不能为空")

        self.messages.append({"role": "user", "content": user_message})
        self._trim_history()

        response = self.client.chat.completions.create(
            model=self.model,
            messages=self.messages,
            temperature=self.temperature,
            max_tokens=self.max_tokens,
        )

        answer = self._extract_content(response).strip()
        if not answer:
            raise RuntimeError("GLM 返回了空响应")

        self.messages.append({"role": "assistant", "content": answer})
        self._trim_history()
        return answer

    def get_history(self, include_system: bool = False) -> List[Dict[str, str]]:
        if include_system:
            return list(self.messages)
        return [m for m in self.messages if m.get("role") != "system"]

    def _trim_history(self):
        system = self.messages[:1]
        non_system = [m for m in self.messages[1:] if m.get("role") in ("user", "assistant")]
        if len(non_system) > self.max_history:
            non_system = non_system[-self.max_history:]
        self.messages = system + non_system

    @staticmethod
    def _extract_content(response) -> str:
        try:
            content = response.choices[0].message.content
        except Exception as e:
            raise RuntimeError(f"GLM 响应解析失败: {e}") from e

        if isinstance(content, str):
            return content
        if isinstance(content, list):
            text_parts = []
            for part in content:
                if isinstance(part, dict) and part.get("type") == "text":
                    text_parts.append(str(part.get("text", "")))
            return "".join(text_parts)
        return str(content or "")

    @staticmethod
    def _build_openai_client(**kwargs):
        try:
            from openai import OpenAI
        except ImportError as e:
            raise RuntimeError("缺少 openai 依赖，请先安装后再使用 AI 对话功能") from e
        return OpenAI(**kwargs)
