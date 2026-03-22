# -*- coding: utf-8 -*-
"""
Anthropic (Claude) API 래퍼 — OpenAI client 인터페이스를 에뮬레이션.

기존 코드의 `client.chat.completions.create(...)` 호출을
수정 없이 그대로 사용할 수 있도록 응답 구조를 맞춤.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


# ── 응답 모방 객체 ──

@dataclass
class _Message:
    content: str = ""
    role: str = "assistant"

@dataclass
class _Choice:
    message: _Message = field(default_factory=_Message)
    index: int = 0

@dataclass
class _ChatResponse:
    choices: list[_Choice] = field(default_factory=list)


# ── Completions 네임스페이스 ──

class _Completions:
    def __init__(self, anthropic_client):
        self._client = anthropic_client

    def create(self, *, model: str, messages: list[dict], **kwargs) -> _ChatResponse:
        # OpenAI messages → Anthropic messages 변환
        system_parts: list[str] = []
        anthropic_msgs: list[dict] = []

        for msg in messages:
            role = msg.get("role", "user")
            raw_content = msg.get("content", "")

            # content가 list(vision 등)인 경우 Anthropic 형식으로 변환
            if isinstance(raw_content, list):
                anthropic_blocks = []
                for block in raw_content:
                    if not isinstance(block, dict):
                        continue
                    btype = block.get("type", "")
                    if btype == "text":
                        anthropic_blocks.append({
                            "type": "text",
                            "text": block.get("text", ""),
                        })
                    elif btype == "image_url":
                        # OpenAI image_url → Anthropic image block
                        url = ""
                        detail = block.get("image_url", {})
                        if isinstance(detail, dict):
                            url = detail.get("url", "")
                        elif isinstance(detail, str):
                            url = detail
                        if url.startswith("data:"):
                            # data:image/jpeg;base64,... 파싱
                            header, _, b64data = url.partition(",")
                            media_type = header.split(";")[0].replace("data:", "")
                            anthropic_blocks.append({
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": b64data,
                                },
                            })

                if role == "system":
                    # system 메시지에서 텍스트만 추출
                    text = "\n".join(
                        b.get("text", "") for b in anthropic_blocks if b.get("type") == "text"
                    )
                    system_parts.append(text)
                else:
                    anthropic_msgs.append({"role": role, "content": anthropic_blocks})
            else:
                text = str(raw_content)
                if role == "system":
                    system_parts.append(text)
                else:
                    anthropic_msgs.append({"role": role, "content": text})

        # 연속 동일 role 메시지 병합 (Anthropic은 user/assistant 교대 필수)
        merged: list[dict] = []
        for m in anthropic_msgs:
            if merged and merged[-1]["role"] == m["role"]:
                prev = merged[-1]["content"]
                curr = m["content"]
                # 둘 다 문자열이면 단순 병합
                if isinstance(prev, str) and isinstance(curr, str):
                    merged[-1]["content"] = prev + "\n\n" + curr
                else:
                    # list 형태로 통일 후 병합
                    def _to_blocks(c):
                        if isinstance(c, list):
                            return list(c)
                        return [{"type": "text", "text": str(c)}]
                    merged[-1]["content"] = _to_blocks(prev) + _to_blocks(curr)
            else:
                merged.append(dict(m))

        # 빈 메시지 방지
        if not merged:
            merged = [{"role": "user", "content": "키워드를 생성해주세요."}]

        # Anthropic API 파라미터 매핑
        api_kwargs: dict[str, Any] = {
            "model": model,
            "messages": merged,
            "max_tokens": kwargs.get("max_tokens", 1024),
        }

        if system_parts:
            api_kwargs["system"] = "\n\n".join(system_parts)

        # Claude API는 temperature와 top_p를 동시에 지정할 수 없음
        # temperature가 있으면 temperature만 사용, 없으면 top_p 사용
        if "temperature" in kwargs:
            api_kwargs["temperature"] = kwargs["temperature"]
        elif "top_p" in kwargs:
            api_kwargs["top_p"] = kwargs["top_p"]

        # response_format={"type": "json_object"} → 프롬프트로 처리
        resp_fmt = kwargs.get("response_format")
        if resp_fmt and resp_fmt.get("type") == "json_object":
            json_hint = "\n\n반드시 유효한 JSON만 출력하세요. 다른 텍스트는 포함하지 마세요."
            if "system" in api_kwargs:
                api_kwargs["system"] += json_hint
            else:
                api_kwargs["system"] = json_hint.strip()

        resp = self._client.messages.create(**api_kwargs)

        # Anthropic 응답 → OpenAI 응답 구조로 변환
        text_out = ""
        if resp.content:
            for block in resp.content:
                if hasattr(block, "text"):
                    text_out += block.text

        return _ChatResponse(
            choices=[_Choice(message=_Message(content=text_out, role="assistant"))]
        )


# ── Chat 네임스페이스 ──

class _Chat:
    def __init__(self, anthropic_client):
        self.completions = _Completions(anthropic_client)


# ── 메인 래퍼 클래스 ──

class AnthropicClientWrapper:
    """
    OpenAI 클라이언트와 동일한 인터페이스를 제공하는 Anthropic 래퍼.

    사용법:
        wrapper = AnthropicClientWrapper(api_key="sk-ant-...")
        resp = wrapper.chat.completions.create(model="claude-haiku-4-5-20251001", messages=[...])
        text = resp.choices[0].message.content
    """

    def __init__(self, api_key: str):
        import anthropic
        self._raw_client = anthropic.Anthropic(api_key=api_key)
        self.chat = _Chat(self._raw_client)
