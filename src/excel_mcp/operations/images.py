"""Pictures placed on a sheet."""

from io import BytesIO

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image as PillowImage
from PIL import UnidentifiedImageError
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.refs import cell_name, parse_cell

PIXELS_PER_CM = 96 / 2.54
EMU_PER_CM = 360_000
MAX_IMAGE_PIXELS = 50_000_000
IMAGE_FORMATS = ("PNG", "JPEG")


class ImageInfo(BaseModel):
    index: int
    anchor: str | None
    width_cm: float | None
    height_cm: float | None


def insert_image(
    sheet: Worksheet, data: bytes, cell: str, width_cm: float | None, height_cm: float | None
) -> None:
    row, column = parse_cell(cell)
    pixel_width, pixel_height = _check_image(data)
    if width_cm is not None and height_cm is not None:
        width, height = width_cm * PIXELS_PER_CM, height_cm * PIXELS_PER_CM
    elif width_cm is not None:
        width = width_cm * PIXELS_PER_CM
        height = width * pixel_height / pixel_width
    elif height_cm is not None:
        height = height_cm * PIXELS_PER_CM
        width = height * pixel_width / pixel_height
    else:
        width, height = pixel_width, pixel_height
    image = Image(BytesIO(data))
    image.width, image.height = round(width), round(height)
    sheet.add_image(image, cell_name(row, column))


def list_images(sheet: Worksheet) -> list[ImageInfo]:
    return [_describe(index, image) for index, image in enumerate(_images(sheet), start=1)]


def delete_image(sheet: Worksheet, index: int) -> ImageInfo:
    images = list_images(sheet)
    if not 1 <= index <= len(images):
        valid = f"1 to {len(images)}" if images else "none: the sheet has no images"
        raise InvalidArgumentError(
            f"Sheet {sheet.title!r} has no image {index}. Valid image indices: {valid}. "
            "describe_sheet lists them."
        )
    del _images(sheet)[index - 1]
    return images[index - 1]


def _images(sheet: Worksheet) -> list[Image]:
    return sheet._images  # pyright: ignore[reportAttributeAccessIssue]


def _check_image(data: bytes) -> tuple[int, int]:
    try:
        with PillowImage.open(BytesIO(data)) as picture:
            image_format, size = picture.format, picture.size
            if image_format not in IMAGE_FORMATS:
                raise InvalidArgumentError(
                    f"The image is {image_format}; only PNG and JPEG images can be inserted."
                )
            if size[0] * size[1] > MAX_IMAGE_PIXELS:
                raise InvalidArgumentError(
                    f"The image is {size[0]}x{size[1]} pixels; the limit is "
                    f"{MAX_IMAGE_PIXELS:,} pixels in total."
                )
            picture.verify()
    except (UnidentifiedImageError, PillowImage.DecompressionBombError, OSError, SyntaxError):
        raise InvalidArgumentError("The file is not a valid PNG or JPEG image.") from None
    return size


def _describe(index: int, image: Image) -> ImageInfo:
    anchor = image.anchor
    if isinstance(anchor, OneCellAnchor):
        ext = anchor.ext
        return ImageInfo(
            index=index,
            anchor=cell_name(anchor._from.row + 1, anchor._from.col + 1),
            width_cm=round(ext.width / EMU_PER_CM, 2),
            height_cm=round(ext.height / EMU_PER_CM, 2),
        )
    if isinstance(anchor, TwoCellAnchor):
        return ImageInfo(
            index=index,
            anchor=cell_name(anchor._from.row + 1, anchor._from.col + 1),
            width_cm=None,
            height_cm=None,
        )
    return ImageInfo(index=index, anchor=None, width_cm=None, height_cm=None)
