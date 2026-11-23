import sys


CREDENTIAL_SERVICE = "RPA-course"


def read_windows_credential(name: str) -> str | None:
    if sys.platform != "win32":
        raise ValueError("Нужен диспетчер учётных данных Windows.")

    from keyring.backends.Windows import WinVaultKeyring
    from keyring.errors import KeyringError
    from win32ctypes.pywin32.pywintypes import error as WindowsCredentialError

    try:
        return WinVaultKeyring().get_password(CREDENTIAL_SERVICE, name)
    except (KeyringError, OSError, WindowsCredentialError):
        raise ValueError("Не удалось прочитать запись из хранилища Windows.") from None
