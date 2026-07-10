#ifndef OPENXLSX_XLXMLFILE_HPP
#define OPENXLSX_XLXMLFILE_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <gsl/gsl>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"
#include "XLXmlSavingDeclaration.hpp"

namespace OpenXLSX
{
    class XLXmlData;
    class XLDocument;
    class XLPackageServices;
    class XLPackagePartFactory;
    class XLSharedStringTable;
    class XLSharedStrings;

    /**
     * @brief The XLUnsupportedElement class provides a stub implementation that can be used as function
     *  parameter or return type for currently unsupported XML features.
     */
    class OPENXLSX_EXPORT XLUnsupportedElement
    {
    public:
        XLUnsupportedElement() {}
        bool        empty() const { return true; }
        std::string summary() const { return "XLUnsupportedElement"; }
    };

    /**
     * @brief The XLXmlFile class provides an interface for derived classes to use.
     * It functions as an ancestor to all classes which are represented by an .xml file in an .xlsx package.
     * @warning The XLXmlFile class is not intended to be instantiated on it's own, but to provide an interface for
     * derived classes. Also, it should not be used polymorphically. For that reason, the destructor is not declared virtual.
     */
    class OPENXLSX_EXPORT XLXmlFile
    {
    public:    // ===== PUBLIC MEMBER FUNCTIONS
        /**
         * @brief Default constructor.
         */
        XLXmlFile() = default;

        /**
         * @brief Method for getting the XML data represented by the object.
         * @param savingDeclaration @optional specify an XML saving declaration to use
         * @return A std::string with the XML data.
         */
        std::string xmlData(XLXmlSavingDeclaration savingDeclaration = XLXmlSavingDeclaration{}) const;

        /**
         * @brief Constructor. Creates an object based on the xmlData input.
         * @param xmlData An XLXmlData object with the XML data to be represented by the object.
         */
        explicit XLXmlFile(XLXmlData* xmlData);

        /**
         * @brief Copy constructor. Default implementation used.
         * @param other The object to copy.
         */
        XLXmlFile(const XLXmlFile& other) = default;

        /**
         * @brief Move constructor. Default implementation used.
         * @param other The object to move.
         */
        XLXmlFile(XLXmlFile&& other) noexcept = default;

        /**
         * @brief Destructor. Default implementation used.
         */
        ~XLXmlFile() = default;

        /**
         * @brief check whether class is linked to a valid XML file
         * @return true if the class should have a link to valid data
         * @return false if accessing any other sheet properties / methods could cause a segmentation fault
         * @note for example, if an XLSheet is created with a default constructor, XLSheetBase::valid() (derived from XLXmlFile) would
         * return false
         */
        bool valid() const { return m_xmlData != nullptr; }

        /**
         * @brief The copy assignment operator. The default implementation has been used.
         * @param other The object to copy.
         * @return A reference to the new object.
         */
        XLXmlFile& operator=(const XLXmlFile& other) = default;

        /**
         * @brief The move assignment operator. The default implementation has been used.
         * @param other The object to move.
         * @return A reference to the new object.
         */
        XLXmlFile& operator=(XLXmlFile&& other) noexcept = default;

        /**
         * @brief This function provides access to the parent XLDocument object.
         * @return A reference to the parent XLDocument object.
         * @note Prefer @ref package() when only archive / XML-part / relationship services are needed.
         */
        XLDocument& parentDoc();

        /**
         * @brief This function provides access to the parent XLDocument object.
         * @return A const reference to the parent XLDocument object.
         * @note Prefer @ref package() when only archive / XML-part / relationship services are needed.
         */
        const XLDocument& parentDoc() const;

        /**
         * @brief Narrow package-services port (archive, managed XML parts, relationships, content types).
         * @details Decouples Drawing / Pivot / similar parts from the full document façade.
         *          Valid only while the parent document remains alive.
         */
        [[nodiscard]] XLPackageServices&       package();
        [[nodiscard]] const XLPackageServices& package() const;

        /**
         * @brief Narrow package-part factory port (chart / drawing / image / slicer / sheet accessories).
         * @details Prefer this over parentDoc() for createChart, sheetDrawing, addImage, createSlicer, etc.
         *          Valid only while the parent document remains alive.
         */
        [[nodiscard]] XLPackagePartFactory&       partFactory();
        [[nodiscard]] const XLPackagePartFactory& partFactory() const;

        /**
         * @brief Read-only shared string table (SST) port.
         * @details Prefer this when only resolving indices to text (or looking up existing strings).
         *          Valid only while the parent document remains alive.
         */
        [[nodiscard]] const XLSharedStringTable& sharedStringTable() const;

        /**
         * @brief Full workbook SST (including getOrCreate / append).
         * @details Convenience forwarding to the parent document; prefer over parentDoc().sharedStrings()
         *          when constructing cells/rows or writing strings. Use @ref sharedStringTable() for
         *          read-only lookups.
         *          Valid only while the parent document remains alive.
         */
        [[nodiscard]] const XLSharedStrings& sharedStrings() const;

        /**
         * @brief This function provides access to the underlying XMLDocument object.
         * @return A reference to the XMLDocument object.
         */
        XMLDocument& xmlDocument();

        /**
         * @brief This function provides access to the underlying XMLDocument object.
         * @return A const reference to the XMLDocument object.
         */
        const XMLDocument& xmlDocument() const;

        /**
         * @brief Retrieve the path of the XML data in the .xlsx zip archive via m_xmlData->getXmlPath
         * @return A std::string with the path. Empty string if m_xmlData is nullptr
         */
        std::string getXmlPath() const;

    protected:    // ===== PROTECTED MEMBER FUNCTIONS
        /**
         * @brief Provide the XML data represented by the object.
         * @param xmlData A std::string_view with the XML data.
         */
        void setXmlData(std::string_view xmlData);

        /**
         * @brief This function returns the relationship ID (the ID used in the XLRelationships objects) for the object.
         * @return A std::string with the ID. Not all spreadsheet objects may have a relationship ID. In those cases an empty string is
         * returned.
         */
        std::string relationshipID() const;

    protected:                         // ===== PROTECTED MEMBER VARIABLES
        XLXmlData* m_xmlData{nullptr}; /**< The underlying XML data object. */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLXMLFILE_HPP
