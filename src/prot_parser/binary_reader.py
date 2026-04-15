from __future__ import annotations


class BinaryReader:
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0

    def remaining(self) -> int:
        return len(self.data) - self.pos

    def read(self, size: int) -> bytes:
        chunk = self.data[self.pos : self.pos + size]
        self.pos += len(chunk)
        return chunk

    def read_u8(self) -> int:
        b = self.read(1)
        return b[0] if b else 0

    def seek(self, offset: int) -> None:
        self.pos = max(0, min(offset, len(self.data)))
