#ifndef OPENXLSX_XLIMAGEPARSER_HPP
#define OPENXLSX_XLIMAGEPARSER_HPP

#include "OpenXLSX-Exports.hpp"
#include <string>
#include <cstdint>

namespace OpenXLSX {

    struct XLImageSize {
        uint32_t width{0};
        uint32_t height{0};
        std::string extension{""};
        bool valid{false};
    };

    /**
     * @brief Parses binary image headers (PNG, JPG, GIF) to extract width and height without external dependencies.
     */
    class OPENXLSX_EXPORT XLImageParser {
    public:
        static XLImageSize parseDimensions(const std::string& path);
    };

} // namespace OpenXLSX

#endif // OPENXLSX_XLIMAGEPARSER_HPP
