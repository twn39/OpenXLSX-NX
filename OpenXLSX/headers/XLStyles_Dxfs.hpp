#ifndef OPENXLSX_XLSTYLES_DXFS_HPP
#define OPENXLSX_XLSTYLES_DXFS_HPP

#ifdef _MSC_VER
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif

#include "XLStyles_Common.hpp"
#include "XLStyles_Borders.hpp"
#include "XLStyles_CellFormats.hpp"
#include "XLStyles_Fills.hpp"
#include "XLStyles_Fonts.hpp"
#include "XLStyles_NumberFormats.hpp"
#include "XLXmlParser.hpp"

#include <memory>
#include <string>
#include <vector>


namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLDxf
    {
        friend class XLDxfs;

    public:
        /**
         * @brief Default constructor. Initializes an empty XLDxf object with a temporary XML document.
         */
        XLDxf();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the dxf item. If no input is provided, a null node is used.
         */
        explicit XLDxf(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLDxf(const XLDxf& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLDxf(XLDxf&& other) noexcept;

        /**
         * @brief Destructor.
         */
        ~XLDxf();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLDxf& operator=(const XLDxf& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLDxf& operator=(XLDxf&& other) noexcept;

        /**
         * @brief Test if this is an empty node
         * @return true if underlying XMLNode is empty
         */
        bool empty() const;

        /**
         * @brief Getter functions, will create empty object on access and can be used to manipulate underlying setters.
         * These are "ensure-and-get" style getters consistent with the rest of the library.
         */
        XLFont         font() const;
        XLNumberFormat numFmt() const;
        XLFill         fill() const;
        XLAlignment    alignment() const;
        XLBorder       border() const;

        /**
         * @brief Check if specific components are present in the DXF.
         */
        bool hasFont() const;
        bool hasNumFmt() const;
        bool hasFill() const;
        bool hasAlignment() const;
        bool hasBorder() const;

        /**
         * @brief Clear specific components from the DXF.
         */
        void clearFont();
        void clearNumFmt();
        void clearFill();
        void clearAlignment();
        void clearBorder();

        /**
         * @brief Unsupported getter
         */
        XLUnsupportedElement extLst() const { return XLUnsupportedElement{}; }

        /**
         * @brief Unsupported setter
         */
        bool setExtLst(XLUnsupportedElement const& newExtLst);

        /**
         * @brief Return a string summary of the differential cell format properties
         * @return string with info about the differential cell format object
         */
        std::string summary() const;

        /**
         * @brief Access the underlying XML node.
         */
        XMLNode node() const { return m_dxfNode; }

    private:
        std::unique_ptr<XMLDocument> m_xmlDocument; /**< An XMLDocument object for standalone use */
        mutable XMLNode              m_dxfNode;     /**< An XMLNode object with the dxf item */
    };

    /**
     * @brief Backward compatibility alias
     */
    using XLDiffCellFormat = XLDxf;

    /**
     * @brief An encapsulation of the XLSX differential cell formats (dxfs)
     */
    class OPENXLSX_EXPORT XLDxfs
    {
    public:
        /**
         * @brief Default constructor.
         */
        XLDxfs();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the dxfs item. If no input is provided, a null node is used.
         */
        explicit XLDxfs(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLDxfs(const XLDxfs& other);

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLDxfs(XLDxfs&& other);

        /**
         * @brief Destructor.
         */
        ~XLDxfs();

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLDxfs& operator=(const XLDxfs& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLDxfs& operator=(XLDxfs&& other) noexcept;

        /**
         * @brief Get the count of differential cell formats
         * @return The amount of entries in the differential cell formats
         */
        size_t count() const;

        /**
         * @brief Get the differential cell format identified by index
         * @return An XLDxf object
         */
        XLDxf dxfByIndex(XLStyleIndex index) const;

        /**
         * @brief Operator overload: allow [] as shortcut access to dxfByIndex
         * @param index The index within the XML sequence
         * @return An XLDxf object
         */
        XLDxf operator[](XLStyleIndex index) const { return dxfByIndex(index); }

        /**
         * @brief Append a new XLDxf, based on copyFrom, and return its index in dxfs node
         * @param copyFrom Can provide an XLDxf to use as template for the new style
         * @param styleEntriesPrefix Prefix the newly created cell style XMLNode with this pugi::node_pcdata text
         * @returns The index of the new differential cell format as used by operator[]
         */
        XLStyleIndex create(XLDxf copyFrom = XLDxf{}, std::string_view styleEntriesPrefix = XLDefaultStyleEntriesPrefix);

    private:
        mutable XMLNode    m_dxfsNode; /**< An XMLNode object with the dxfs item */
        std::vector<XLDxf> m_dxfs;
    };

    /**
     * @brief Backward compatibility alias
     */
    using XLDiffCellFormats = XLDxfs;

}    // namespace OpenXLSX

#ifdef _MSC_VER
#    pragma warning(pop)
#endif

#endif    // OPENXLSX_XLSTYLES_DXFS_HPP
