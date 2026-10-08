"""MS-OVBA compression ([MS-OVBA] section 2.4.1), the inverse of ``ovba.decompress``."""

_CHUNK_SIZE = 4096
_MIN_MATCH = 3
_MAX_CANDIDATES = 32


def compress(data: bytes) -> bytes:
    output = bytearray([0x01])
    for start in range(0, len(data), _CHUNK_SIZE):
        output += _compress_chunk(data[start : start + _CHUNK_SIZE])
    return bytes(output)


def _compress_chunk(chunk: bytes) -> bytes:
    body = bytearray()
    seen: dict[bytes, list[int]] = {}
    position = 0
    while position < len(chunk):
        flags_index = len(body)
        body.append(0)
        for bit in range(8):
            if position >= len(chunk):
                break
            offset, length = _longest_match(chunk, position, seen)
            if length:
                bit_count = max((position - 1).bit_length(), 4)
                token = ((offset - 1) << (16 - bit_count)) | (length - _MIN_MATCH)
                body[flags_index] |= 1 << bit
                body += token.to_bytes(2, "little")
            else:
                body.append(chunk[position])
                length = 1
            for index in range(position, min(position + length, len(chunk) - _MIN_MATCH + 1)):
                seen.setdefault(chunk[index : index + _MIN_MATCH], []).append(index)
            position += length
    if len(body) < _CHUNK_SIZE:
        return (0xB000 | (len(body) - 1)).to_bytes(2, "little") + body
    # Incompressible chunks are stored as is, padded to the full chunk size.
    return (0x3000 | (_CHUNK_SIZE - 1)).to_bytes(2, "little") + chunk.ljust(_CHUNK_SIZE, b"\0")


def _longest_match(chunk: bytes, position: int, seen: dict[bytes, list[int]]) -> tuple[int, int]:
    """Find the longest earlier occurrence of the text at ``position``: (offset, length)."""
    bit_count = max((position - 1).bit_length(), 4)
    limit = min((0xFFFF >> bit_count) + _MIN_MATCH, len(chunk) - position)
    best_offset = best_length = 0
    for start in reversed(seen.get(chunk[position : position + _MIN_MATCH], [])[-_MAX_CANDIDATES:]):
        length = _MIN_MATCH
        while length < limit and chunk[start + length] == chunk[position + length]:
            length += 1
        if length > best_length:
            best_offset, best_length = position - start, length
    return best_offset, best_length
