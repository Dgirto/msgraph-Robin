from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional


@dataclass
class Attachment:
    filename: str
    content_type: str
    content_bytes: str  # base64


@dataclass
class Email:
    id: str
    subject: str
    from_address: str
    to: List[str]
    body: str
    received_at: datetime
    attachments: List[Attachment] = field(default_factory=list)


@dataclass
class FileObject:
    id: str
    name: str
    path: Optional[str]
    size: int
    mime_type: Optional[str]
    drive_id: Optional[str]