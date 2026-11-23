from functools import cached_property
from http.client import HTTPException
import json
import ssl
from time import sleep, time
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import HTTPRedirectHandler, HTTPSHandler, Request, build_opener
from uuid import uuid4

from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlsplit

from credentials import read_windows_credential


OAUTH_URL = "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
TIMEOUT = 60
MAX_ATTEMPTS = 3
RETRY_INTERVAL = 30


@dataclass(frozen=True)
class GigaChatConfig:
    scope: str
    model: str
    base_url: str
    verify_ssl_certs: bool
    ca_bundle_file: Path | None = None

    def __post_init__(self):
        url = urlsplit(self.base_url)
        if url.scheme != "https" or not url.hostname or url.username or url.query or url.fragment:
            raise ValueError("Укажите адрес HTTPS для GigaChat.")
        if self.verify_ssl_certs is not True:
            raise ValueError("Для GigaChat требуется проверка TLS.")


class GigaChatError(ValueError):
    """Ошибка вызова GigaChat без заголовков и содержимого запросов."""


class GigaChatSetupError(ValueError):
    """Для продолжения нужно исправить настройки или учётные данные."""


class GigaChatTimeout(GigaChatError):
    pass


class InvalidModelResponse(ValueError):
    """Модель вернула ответ, который не прошёл проверку."""


class NoRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class GigaChatClient:
    def __init__(self, config: GigaChatConfig):
        self.config = config
        self._token = ""
        self._token_expires = 0.0
        self._unavailable = None

    @cached_property
    def _opener(self):
        try:
            context = ssl.create_default_context()
            if self.config.ca_bundle_file is not None:
                context.load_verify_locations(cafile=str(self.config.ca_bundle_file))
        except OSError:
            raise GigaChatSetupError("Не удалось загрузить доверенные сертификаты GigaChat.") from None
        return build_opener(HTTPSHandler(context=context), NoRedirects())

    def _request_json(self, url: str, data: bytes, headers: dict[str, str]):
        opener = self._opener
        oauth = url == OAUTH_URL
        stage = "OAuth" if oauth else "сопоставление"
        request_headers = {"Accept": "application/json", **headers}
        for attempt in range(1, MAX_ATTEMPTS + 1):
            # OAuth требует отдельный идентификатор каждого HTTP-запроса, включая повтор.
            if oauth:
                request_headers["RqUID"] = str(uuid4())
            request = Request(url, data=data, headers=request_headers, method="POST")
            try:
                with opener.open(request, timeout=TIMEOUT) as response:
                    return json.load(response)
            except HTTPError as error:
                code = error.code
                error.close()
                if code == 408:
                    raise GigaChatTimeout(f"Истекло время ожидания GigaChat ({stage}).") from None
                if code != 429 and not 500 <= code < 600:
                    rejected = "авторизацию" if code in {401, 403} else "запрос"
                    raise GigaChatSetupError(f"GigaChat отклонил {rejected} ({stage}, HTTP {code}).") from None
            except (OSError, HTTPException) as error:
                reason = error.reason if isinstance(error, URLError) else error
                if isinstance(reason, TimeoutError):
                    raise GigaChatTimeout(f"Истекло время ожидания GigaChat ({stage}).") from None
                if isinstance(reason, ssl.SSLCertVerificationError):
                    raise GigaChatSetupError("GigaChat не прошёл проверку TLS. Проверьте сертификат Минцифры.") from None
            except (ValueError, RecursionError):
                if oauth:
                    raise GigaChatError("GigaChat вернул некорректный ответ OAuth.") from None
                raise InvalidModelResponse from None

            if attempt < MAX_ATTEMPTS:
                sleep(RETRY_INTERVAL)
        raise GigaChatError(f"GigaChat недоступен после трёх попыток ({stage}).")

    def _get_token(self) -> str:
        if self._token and time() + 30 < self._token_expires:
            return self._token

        try:
            key = (read_windows_credential("GIGACHAT_CREDENTIALS") or "").strip()
        except (ValueError, ImportError):
            raise GigaChatSetupError("Не удалось прочитать ключ GigaChat из хранилища Windows.") from None
        if not key:
            raise GigaChatSetupError("Сохраните GIGACHAT_CREDENTIALS в хранилище Windows. См. README.md.")
        if not key.isascii() or any(char.isspace() for char in key):
            raise GigaChatSetupError("В хранилище Windows нужен ключ GigaChat без префикса Basic.")

        try:
            result = self._request_json(
                OAUTH_URL, urlencode({"scope": self.config.scope}).encode("ascii"),
                {"Authorization": f"Basic {key}", "Content-Type": "application/x-www-form-urlencoded"},
            )
        except GigaChatTimeout as error:
            raise GigaChatError(str(error)) from None
        if not isinstance(result, dict):
            raise GigaChatError("GigaChat вернул некорректный ответ OAuth.")
        token, expires = result.get("access_token"), result.get("expires_at")
        if not isinstance(token, str) or not token or not token.isascii() or any(c.isspace() for c in token):
            raise GigaChatError("GigaChat не вернул корректный токен доступа.")
        if type(expires) is not int or expires <= 0:
            raise GigaChatError("GigaChat не вернул срок действия токена.")
        # Сервис возвращает Unix timestamp. Поддерживаем секунды и миллисекунды.
        self._token_expires = expires / 1000 if expires >= 100_000_000_000 else expires
        self._token = token
        return token

    def complete(self, payload: dict) -> str:
        if self._unavailable:
            raise GigaChatError(self._unavailable)
        try:
            result = self._request_json(
                f"{self.config.base_url.rstrip('/')}/chat/completions",
                json.dumps({"model": self.config.model, **payload}, ensure_ascii=False).encode("utf-8"),
                {"Authorization": f"Bearer {self._get_token()}", "Content-Type": "application/json"},
            )
        except (GigaChatSetupError, GigaChatError) as error:
            if not isinstance(error, GigaChatTimeout):
                self._unavailable = str(error)
            raise
        try:
            choice = result["choices"][0]
            content = choice["message"]["content"]
            if choice["finish_reason"] != "stop" or not isinstance(content, str) or not content.strip():
                raise InvalidModelResponse
            return content
        except (KeyError, TypeError, IndexError):
            raise InvalidModelResponse from None
