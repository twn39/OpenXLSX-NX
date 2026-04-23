#ifndef OPENXLSX_XLXMLDATA_HPP
#define OPENXLSX_XLXMLDATA_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <cstdlib>
#include <memory>
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLContentTypes.hpp"
#include "XLXmlParser.hpp"
#include "XLXmlSavingDeclaration.hpp"

namespace OpenXLSX
{
    /**
     * @brief Auto-managed malloc memory for Zero-Copy serialization
     */
    struct XLAllocatedMemory {
        void* data = nullptr;
        size_t size = 0;

        ~XLAllocatedMemory() { if (data) std::free(data); }
        XLAllocatedMemory() = default;
        
        // Disable copy
        XLAllocatedMemory(const XLAllocatedMemory&) = delete;
        XLAllocatedMemory& operator=(const XLAllocatedMemory&) = delete;
        
        // Enable move
        XLAllocatedMemory(XLAllocatedMemory&& other) noexcept : data(other.data), size(other.size) {
            other.data = nullptr;
            other.size = 0;
        }
        XLAllocatedMemory& operator=(XLAllocatedMemory&& other) noexcept {
            if (this != &other) {
                if (data) std::free(data);
                data = other.data;
                size = other.size;
                other.data = nullptr;
                other.size = 0;
            }
            return *this;
        }

        // Release ownership
        void* release() {
            void* ret = data;
            data = nullptr;
            return ret;
        }
    };

    /**
     * @brief The XLXmlData class encapsulates the properties and behaviour of the .xml files in an .xlsx file zip
     * package. Objects of the XLXmlData type are intended to be stored centrally in an XLDocument object, from where
     * they can be retrieved by other objects that encapsulates the behaviour of Excel elements, such as XLWorkbook
     * and XLWorksheet.
     */
    class OPENXLSX_EXPORT XLXmlData final
    {
    public:
        // ===== PUBLIC MEMBER FUNCTIONS ===== //

        /**
         * @brief Default constructor. All member variables are default constructed. Except for
         * the raw XML data, none of the member variables can be modified after construction. Hence, objects created
         * using the default constructor can only serve as null objects and targets for the move assignemnt operator.
         */
        XLXmlData() = default;

        /**
         * @brief Constructor. This constructor creates objects with the given parameters. the xmlId and the xmlType
         * parameters have default values. These are only useful for relationship (.rels) files and the
         * [Content_Types].xml file located in the root directory of the zip package.
         * @param parentDoc A pointer to the parent XLDocument object.
         * @param xmlPath A std::string with the file path in zip package.
         * @param xmlId A std::string with the relationship ID of the file (used in the XLRelationships class)
         * @param xmlType The type of object the XML file represents, e.g. XLWorkbook or XLWorksheet.
         */
        XLXmlData(XLDocument*        parentDoc,
                  const std::string& xmlPath,
                  const std::string& xmlId   = "",
                  XLContentType      xmlType = XLContentType::Unknown);

        /**
         * @brief Default destructor. The XLXmlData does not manage any dynamically allocated resources, so a default
         * destructor will suffice.
         */
        ~XLXmlData();

        /**
         * @brief check whether class is linked to a valid XML document
         * @return true if the class should have a link to valid data
         * @return false if accessing any other properties / methods could cause a segmentation fault
         */
        bool valid() const { return m_xmlDoc != nullptr; }

        /**
         * @brief Copy constructor. The m_xmlDoc data member is a XMLDocument object, which is non-copyable. Hence,
         * the XLXmlData objects have a explicitly deleted copy constructor.
         * @param other
         */
        XLXmlData(const XLXmlData& other) = delete;

        /**
         * @brief Move constructor. All data members are trivially movable. Hence an explicitly defaulted move
         * constructor is sufficient.
         * @param other
         */
        XLXmlData(XLXmlData&& other) noexcept = default;

        /**
         * @brief Copy assignment operator. The m_xmlDoc data member is a XMLDocument object, which is non-copyable.
         * Hence, the XLXmlData objects have a explicitly deleted copy assignment operator.
         */
        XLXmlData& operator=(const XLXmlData& other) = delete;

        /**
         * @brief Move assignment operator. All data members are trivially movable. Hence an explicitly defaulted move
         * constructor is sufficient.
         * @param other the XLXmlData object to be moved from.
         * @return A reference to the moved-to object.
         */
        XLXmlData& operator=(XLXmlData&& other) noexcept = default;

        /**
         * @brief Set the raw data for the underlying XML document. Being able to set the XML data directly is useful
         * when creating a new file using a XML file template. E.g., when creating a new worksheet, the XML code for
         * a minimum viable XLWorksheet object can be added using this function.
         * @param data A std::string with the raw XML text.
         */
        void setRawData(const std::string& data);

        /**
         * @brief Get the raw data for the underlying XML document. This function will retrieve the raw XML text data
         * from the underlying XMLDocument object. This will mainly be used when saving data to the .xlsx package
         * using the save function in the XLDocument class.
         * @param savingDeclaration @optional specify an XML saving declaration to use
         * @return A std::string with the raw XML text data.
         */
        std::string getRawData(XLXmlSavingDeclaration savingDeclaration = XLXmlSavingDeclaration{}) const;

        /**
         * @brief Get the raw data allocated directly using malloc. Intended for zero-copy serialization
         * straight into libzip buffers, avoiding unnecessary std::string and string stream allocations.
         * @param savingDeclaration @optional specify an XML saving declaration to use
         * @return An XLAllocatedMemory object holding ownership of the buffer and its size.
         */
        XLAllocatedMemory getRawAllocatedData(XLXmlSavingDeclaration savingDeclaration = XLXmlSavingDeclaration{}) const;

        /**
         * @brief Access the parent XLDocument object.
         * @return A pointer to the parent XLDocument object.
         */
        XLDocument* getParentDoc();

        /**
         * @brief Access the parent XLDocument object.
         * @return A const pointer to the parent XLDocument object.
         */
        const XLDocument* getParentDoc() const;

        /**
         * @brief Retrieve the path of the XML data in the .xlsx zip archive.
         * @return A std::string with the path.
         */
        std::string getXmlPath() const;

        /**
         * @brief Retrieve the relationship ID of the XML file.
         * @return A std::string with the relationship ID.
         */
        std::string getXmlID() const;

        /**
         * @brief Retrieve the type represented by the XML data.
         * @return A XLContentType getValue representing the type.
         */
        XLContentType getXmlType() const;

        /**
         * @brief Access the underlying XMLDocument object.
         * @return A pointer to the XMLDocument object.
         */
        XMLDocument* getXmlDocument();

        /**
         * @brief Access the underlying XMLDocument object.
         * @return A const pointer to the XMLDocument object.
         */
        const XMLDocument* getXmlDocument() const;

        /**
         * @brief Test whether there is an XML file linked to this object
         * @return true if there is no underlying XML file, otherwise false
         */
        bool empty() const;

    private:
        // ===== PRIVATE MEMBER VARIABLES ===== //

        XLDocument*                          m_parentDoc{}; /**< A pointer to the parent XLDocument object. >*/
        std::string                          m_xmlPath{};   /**< The path of the XML data in the .xlsx zip archive. >*/
        std::string                          m_xmlID{};     /**< The relationship ID of the XML data. >*/
        XLContentType                        m_xmlType{};   /**< The type represented by the XML data. >*/
        mutable std::unique_ptr<XMLDocument> m_xmlDoc;      /**< The underlying XMLDocument object. >*/
    public:
        bool        m_isStreamed{false};
        std::string m_streamFilePath;
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLXMLDATA_HPP
