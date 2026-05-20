#ifndef OPENXLSX_XLSLICER_HPP
#define OPENXLSX_XLSLICER_HPP

#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include <string>
#include <string_view>

namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLSlicer final : public XLXmlFile
    {
    public:
        /**
         * @brief Default constructor.
         */
        XLSlicer() : XLXmlFile(nullptr) {}

        /**
         * @brief Constructor that wraps an XLXmlData pointer.
         */
        explicit XLSlicer(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

        ~XLSlicer() = default;
        XLSlicer(const XLSlicer& other) = default;
        XLSlicer(XLSlicer&& other) noexcept = default;
        XLSlicer& operator=(const XLSlicer& other) = default;
        XLSlicer& operator=(XLSlicer&& other) noexcept = default;

        /**
         * @brief Get the name of the slicer.
         */
        std::string name() const;

        /**
         * @brief Set the name of the slicer.
         */
        void setName(std::string_view name);

        /**
         * @brief Get the caption of the slicer.
         */
        std::string caption() const;

        /**
         * @brief Set the caption of the slicer.
         */
        void setCaption(std::string_view caption);

        /**
         * @brief Get the style name of the slicer.
         */
        std::string slicerStyle() const;

        /**
         * @brief Set the style name of the slicer (e.g. "SlicerStyleLight1").
         */
        void setSlicerStyle(std::string_view styleName);

        /**
         * @brief Check if the slicer caption is shown.
         */
        bool showCaption() const;

        /**
         * @brief Set whether to show the slicer caption.
         */
        void setShowCaption(bool show);

        /**
         * @brief Get the associated slicer cache name.
         */
        std::string cache() const;

        /**
         * @brief Set the associated slicer cache name.
         */
        void setCache(std::string_view cacheName);
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLSLICER_HPP
