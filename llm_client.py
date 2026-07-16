from __future__ import annotations

import os

try:
    from anthropic import Anthropic
except ModuleNotFoundError:
    # ローカル要約のみの実行では anthropic パッケージ不要。
    # API利用時（--api-summary 等）に get_client() が明示エラーを出す。
    Anthropic = None  # type: ignore[assignment]

DEFAULT_MAX_TOKENS = 8192
WEB_SEARCH_TOOL = {"type": "web_search_20250305", "name": "web_search"}


def get_client() -> "Anthropic":
    if Anthropic is None:
        raise RuntimeError(
            "anthropic パッケージが見つかりません。pip install -r requirements.txt を実行してください。"
        )
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise RuntimeError(
            "Anthropic client を作成できませんでした。ANTHROPIC_API_KEY を確認してください。"
        )
    return Anthropic()


def create_response(
    client: Anthropic,
    model: str,
    prompt: str,
    *,
    max_tokens: int = DEFAULT_MAX_TOKENS,
    timeout: float | None = None,
    web_search: bool = False,
) -> str:
    """1つのuserメッセージとしてprompt を送り、テキスト出力を結合して返す。

    旧OpenAI Responses APIの `response.output_text` と同じ使い勝手にすることで、
    呼び出し側のロジックを変えずにクライアントだけ差し替えられるようにしている。
    """
    kwargs: dict = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": prompt}],
    }
    if timeout is not None:
        kwargs["timeout"] = timeout

    if web_search:
        try:
            return _extract_text(client.messages.create(tools=[WEB_SEARCH_TOOL], **kwargs))
        except Exception:
            pass  # web検索ツールが使えない場合はツールなしで再試行する

    return _extract_text(client.messages.create(**kwargs))


def _extract_text(message) -> str:
    return "".join(
        block.text for block in message.content if getattr(block, "type", None) == "text"
    ).strip()
