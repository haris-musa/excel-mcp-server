"""Pictures placed on a sheet."""

from io import BytesIO

from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image as PillowImage
from PIL import UnidentifiedImageError
from pydantic import BaseModel

from excel_mcp.errors import InvalidArgumentError
from excel_mcp.operations import drawings
from excel_mcp.operations.chart_index import name_shapes, shape_names
from excel_mcp.package.shape_names import name_of, set_name
from excel_mcp.refs import cell_name, parse_cell

PIXELS_PER_CM = 96 / 2.54
EMU_PER_CM = drawings.EMU_PER_CM
MAX_IMAGE_PIXELS = 50_000_000
IMAGE_FORMATS = ("PNG", "JPEG")


class ImageInfo(BaseModel):
    name: str
    range: str | None = None
    width_cm: float | None = None
    height_cm: float | None = None


def insert_image(
    sheet: Worksheet,
    data: bytes,
    at: str,
    width_cm: float | None,
    height_cm: float | None,
    name: str | None,
) -> ImageInfo:
    row, column = parse_cell(at)
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
    taken = shape_names(sheet)
    image = Image(BytesIO(data))
    image.width, image.height = round(width), round(height)
    set_name(
        image,
        drawings.check_name(name, taken)
        if name is not None
        else drawings.free_name("Picture", taken),
    )
    sheet.add_image(image, cell_name(row, column))
    return _describe(sheet, image)


def list_images(sheet: Worksheet) -> list[ImageInfo]:
    name_shapes(sheet)
    return [_describe(sheet, image) for image in _images(sheet)]


def delete_image(sheet: Worksheet, name: str) -> ImageInfo:
    images = list_images(sheet)
    index = drawings.find([image.name for image in images], name, "image", sheet)
    del _images(sheet)[index]
    return images[index]


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


def _describe(sheet: Worksheet, image: Image) -> ImageInfo:
    anchor = image.anchor
    name = name_of(image) or ""
    if isinstance(anchor, str):
        row, col = parse_cell(anchor)
        width, height = image.width * drawings.EMU_PER_PIXEL, image.height * drawings.EMU_PER_PIXEL
        area = drawings.extent(sheet, row, col, width, height)
        return ImageInfo(
            name=name,
            range=area,
            width_cm=round(width / EMU_PER_CM, 2),
            height_cm=round(height / EMU_PER_CM, 2),
        )
    area = drawings.covered(sheet, anchor)
    if isinstance(anchor, OneCellAnchor):
        return ImageInfo(
            name=name,
            range=area,
            width_cm=round(anchor.ext.width / EMU_PER_CM, 2),
            height_cm=round(anchor.ext.height / EMU_PER_CM, 2),
        )
    return ImageInfo(name=name, range=area, width_cm=None, height_cm=None)
