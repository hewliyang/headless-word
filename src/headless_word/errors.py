class HeadlessWordError(Exception):
    pass


class DaemonError(HeadlessWordError):
    pass


class LibreOfficeNotFoundError(HeadlessWordError):
    pass


class SessionError(HeadlessWordError):
    pass


class ToolError(HeadlessWordError):
    pass
