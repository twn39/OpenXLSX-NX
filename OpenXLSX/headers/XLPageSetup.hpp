#ifndef OPENXLSX_XLPAGESETUP_HPP
#define OPENXLSX_XLPAGESETUP_HPP

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"
#include "XLConstants.hpp"

namespace OpenXLSX
{
    /**
     * @brief Enum defining page orientation.
     */
    enum class XLPageOrientation { Default, Portrait, Landscape };

    /**
     * @brief A class representing the page margins of a worksheet.
     */
    class OPENXLSX_EXPORT XLPageMargins
    {
    public:
        XLPageMargins() = default;
        explicit XLPageMargins(const XMLNode& node);

        double left() const;
        XLPageMargins& setLeft(double value);
        XLPageMargins& setLeft(XLDistance value);

        double right() const;
        XLPageMargins& setRight(double value);
        XLPageMargins& setRight(XLDistance value);

        double top() const;
        XLPageMargins& setTop(double value);
        XLPageMargins& setTop(XLDistance value);

        double bottom() const;
        XLPageMargins& setBottom(double value);
        XLPageMargins& setBottom(XLDistance value);

        double header() const;
        XLPageMargins& setHeader(double value);
        XLPageMargins& setHeader(XLDistance value);

        double footer() const;
        XLPageMargins& setFooter(double value);
        XLPageMargins& setFooter(XLDistance value);

    private:
        XMLNode m_node;
    };

    /**
     * @brief A class representing the print options of a worksheet.
     */
    class OPENXLSX_EXPORT XLPrintOptions
    {
    public:
        XLPrintOptions() = default;
        explicit XLPrintOptions(const XMLNode& node);

        bool gridLines() const;
        void setGridLines(bool value);

        bool headings() const;
        void setHeadings(bool value);

        bool horizontalCentered() const;
        void setHorizontalCentered(bool value);

        bool verticalCentered() const;
        void setVerticalCentered(bool value);

    private:
        XMLNode m_node;
    };

    /**
     * @brief A class representing the page setup of a worksheet.
     */
    /**
     * @brief A class representing the header and footer of a worksheet.
     */
    class OPENXLSX_EXPORT XLHeaderFooter
    {
    public:
        XLHeaderFooter() = default;
        explicit XLHeaderFooter(const XMLNode& node);

        [[nodiscard]] bool differentFirst() const;
        void setDifferentFirst(bool value);

        [[nodiscard]] bool differentOddEven() const;
        void setDifferentOddEven(bool value);

        [[nodiscard]] bool scaleWithDoc() const;
        void setScaleWithDoc(bool value);

        [[nodiscard]] bool alignWithMargins() const;
        void setAlignWithMargins(bool value);

        [[nodiscard]] std::string oddHeader() const;
        void setOddHeader(std::string_view value);

        [[nodiscard]] std::string oddFooter() const;
        void setOddFooter(std::string_view value);

        [[nodiscard]] std::string evenHeader() const;
        void setEvenHeader(std::string_view value);

        [[nodiscard]] std::string evenFooter() const;
        void setEvenFooter(std::string_view value);

        [[nodiscard]] std::string firstHeader() const;
        void setFirstHeader(std::string_view value);

        [[nodiscard]] std::string firstFooter() const;
        void setFirstFooter(std::string_view value);

    private:
        XMLNode m_node;
    };


    class OPENXLSX_EXPORT XLPageSetup
    {
    public:
        XLPageSetup() = default;
        explicit XLPageSetup(const XMLNode& node);

        uint32_t paperSize() const;
        XLPageSetup& setPaperSize(uint32_t value);

        XLPageOrientation orientation() const;
        XLPageSetup& setOrientation(XLPageOrientation value);

        uint32_t scale() const;
        XLPageSetup& setScale(uint32_t value);

        uint32_t fitToWidth() const;
        XLPageSetup& setFitToWidth(uint32_t value);


        uint32_t fitToHeight() const;
        XLPageSetup& setFitToHeight(uint32_t value);

        [[nodiscard]] std::string pageOrder() const;
        XLPageSetup& setPageOrder(std::string_view value);

        [[nodiscard]] bool useFirstPageNumber() const;
        XLPageSetup& setUseFirstPageNumber(bool value);

        [[nodiscard]] uint32_t firstPageNumber() const;
        XLPageSetup& setFirstPageNumber(uint32_t value);

        bool blackAndWhite() const;
        XLPageSetup& setBlackAndWhite(bool value);

    private:
        XMLNode m_node;
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLPAGESETUP_HPP
