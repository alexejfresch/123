"""Официальный Sheets API, OAuth только для чтения, токены в системном сейфе."""
import json
import re
import uuid
from pathlib import Path

import keyring
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google_auth_httplib2 import AuthorizedHttp
import httplib2

from .excel import validate_headers

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SERVICE = "UIK-Reconcile"


def spreadsheet_id(value):
    value = value.strip()
    match = re.search(r"/spreadsheets/d/([\w-]+)", value)
    result = match.group(1) if match else value
    if not re.fullmatch(r"[A-Za-z0-9_-]+", result):
        raise ValueError("Укажите ID или ссылку Google Таблицы")
    return result


class GoogleSource:
    def __init__(self, root, account):
        self.root = Path(root)
        self.account = account

    def authorize(self, client_file):
        # Никакого запасного текстового хранилища: отсутствие сейфа — явная ошибка.
        try:
            keyring.get_password(SERVICE, self.account)
        except Exception as exc:
            raise RuntimeError("Недоступно защищённое хранилище ОС. На Windows нужен Диспетчер учётных данных.") from exc
        flow = InstalledAppFlow.from_client_secrets_file(str(client_file), SCOPES)
        credentials = flow.run_local_server(
            port=0, timeout_seconds=180, access_type="offline", prompt="consent",
            authorization_prompt_message="Откройте браузер для авторизации Google.",
            success_message="Авторизация завершена. Вернитесь в программу.",
        )
        keyring.set_password(SERVICE, self.account, credentials.to_json())

    def _api(self):
        try:
            token = keyring.get_password(SERVICE, self.account)
            if not token:
                raise RuntimeError("Выполните вход Google в разделе «Источники данных»")
            credentials = Credentials.from_authorized_user_info(json.loads(token), SCOPES)
            if credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
                keyring.set_password(SERVICE, self.account, credentials.to_json())
            return build("sheets", "v4", http=AuthorizedHttp(credentials, http=httplib2.Http(timeout=30)),
                         cache_discovery=False, static_discovery=True)
        except RuntimeError:
            raise
        except Exception as exc:
            raise RuntimeError("Не удалось получить авторизацию Google. Повторите вход.") from exc

    def sheets(self, document):
        api = self._api()
        try:
            result = api.spreadsheets().get(spreadsheetId=spreadsheet_id(document),
                                          fields="sheets.properties.title").execute(num_retries=1)
            return [s["properties"]["title"] for s in result.get("sheets", [])]
        except Exception as exc:
            raise RuntimeError("Не удалось прочитать список листов Google. Проверьте доступ, ID и интернет.") from exc
        finally:
            api.close()

    def read(self, document, sheet):
        api = self._api()
        try:
            title = "'" + sheet.replace("'", "''") + "'"
            values = api.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id(document), range=title,
                valueRenderOption="FORMATTED_VALUE", dateTimeRenderOption="FORMATTED_STRING",
            ).execute(num_retries=1).get("values", [])
        except Exception as exc:
            raise RuntimeError("Ошибка чтения Google. Сохранённые данные доступны; проверьте подключение и права.") from exc
        finally:
            api.close()
        if not values:
            return [], []
        width = max(map(len, values))
        headers = [str(v).strip() for v in values[0]]
        if len(headers) < width:
            raise ValueError("В Google есть данные в столбцах без заголовка")
        validate_headers(headers)
        # Пустые промежуточные строки сохраняют исходные номера остальных записей.
        rows = [dict(zip(headers, row + [None] * (width - len(row)))) for row in values[1:]]
        return headers, rows


def new_account():
    return str(uuid.uuid4())
