/*
 * OpenXLSX-NX - Modern High-Performance C++ Excel Library
 *
 * Original work Copyright (c) 2018-2025, Kenneth Troldal Balslev and OpenXLSX contributors
 * Modified work Copyright (c) 2026, Curry Tang and OpenXLSX-NX contributors
 *
 * Distributed under the BSD 3-Clause License.
 * See LICENSE.md in the project root or https://opensource.org/licenses/BSD-3-Clause
 */

#ifndef OPENXLSX_XLIMAGEOPTIONS_HPP
#define OPENXLSX_XLIMAGEOPTIONS_HPP

#include <cstdint>
#include <string>

namespace OpenXLSX
{

    /**
     * @brief Determines how an image behaves when cells are resized or inserted/deleted.
     */
    enum class XLImagePositioning {
        OneCell, /**< Image moves with the top-left cell but does not resize. */
        TwoCell, /**< Image stretches and moves with both top-left and bottom-right cells. */
        Absolute /**< Image stays exactly at a fixed absolute coordinate on the sheet. */
    };

    /**
     * @brief Options for inserting an image into a worksheet.
     */
    struct XLImageOptions
    {
        double scaleX{1.0}; /**< Horizontal scaling factor. 1.0 = 100%. */
        double scaleY{1.0}; /**< Vertical scaling factor. 1.0 = 100%. */

        int32_t offsetX{0}; /**< Horizontal offset in pixels from the anchor cell's left edge. */
        int32_t offsetY{0}; /**< Vertical offset in pixels from the anchor cell's top edge. */

        XLImagePositioning positioning{XLImagePositioning::OneCell}; /**< How the image anchors to cells. */

        std::string bottomRightCell{
            ""}; /**< Bottom-right cell reference (e.g., "D10"). Used if positioning is TwoCell. If empty, calculated automatically. */

        bool printWithSheet{true}; /**< Whether the image should be printed with the worksheet. */
        bool locked{false};        /**< Whether the image is locked (requires sheet protection to be active). */
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLIMAGEOPTIONS_HPP
