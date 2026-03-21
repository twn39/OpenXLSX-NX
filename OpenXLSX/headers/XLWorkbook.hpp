#ifndef OPENXLSX_XLWORKBOOK_HPP
#define OPENXLSX_XLWORKBOOK_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <optional>
#include <ostream>    // std::basic_ostream
#include <string>
#include <string_view>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    class XLSheet;
    class XLWorksheet;
    class XLChartsheet;

    /**
     * @brief Enumeration of the different types of sheets.
     */
    enum class XLSheetType { Worksheet, Chartsheet, Dialogsheet, Macrosheet };

    /**
     * @brief A class representing a single defined name in a workbook.
     */
    class OPENXLSX_EXPORT XLDefinedName
    {
    public:
        XLDefinedName() = default;
        explicit XLDefinedName(const XMLNode& node);

        std::string name() const;
        void        setName(std::string_view name);

        std::string refersTo() const;
        void        setRefersTo(std::string_view formula);

        std::optional<uint32_t> localSheetId() const;
        void                    setLocalSheetId(uint32_t id);

        bool hidden() const;
        void setHidden(bool hidden);

        std::string comment() const;
        void        setComment(std::string_view comment);

        bool valid() const { return !m_node.empty(); }

    private:
        XMLNode m_node;
    };

    /**
     * @brief A class representing the collection of defined names in a workbook.
     */
    class OPENXLSX_EXPORT XLDefinedNames
    {
    public:
        XLDefinedNames() = default;
        explicit XLDefinedNames(const XMLNode& node);

        XLDefinedName append(std::string_view name, std::string_view formula, std::optional<uint32_t> localSheetId = std::nullopt);
        void          remove(std::string_view name, std::optional<uint32_t> localSheetId = std::nullopt);
        XLDefinedName get(std::string_view name, std::optional<uint32_t> localSheetId = std::nullopt) const;
        std::vector<XLDefinedName> all() const;
        bool                       exists(std::string_view name, std::optional<uint32_t> localSheetId = std::nullopt) const;
        size_t                     count() const;

    private:
        XMLNode m_node;
    };

    /**
     * @brief This class encapsulates the concept of a Workbook. It provides access to the individual sheets
     * (worksheets or chartsheets), as well as functionality for adding, deleting, moving and renaming sheets.
     */
    class OPENXLSX_EXPORT XLWorkbook : public XLXmlFile
    {
        friend class XLSheet;
        friend class XLDocument;

    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief Default constructor. Creates an empty ('null') XLWorkbook object.
         */
        XLWorkbook() = default;

        /**
         * @brief Constructor. Takes a pointer to an XLXmlData object (stored in the parent XLDocument object).
         * @param xmlData A pointer to the underlying XLXmlData object, which holds the XML data.
         * @note Do not create an XLWorkbook object directly. Access via XLDocument::workbook().
         */
        explicit XLWorkbook(XLXmlData* xmlData);

        /**
         * @brief Copy Constructor.
         * @param other The XLWorkbook object to be copied.
         */
        XLWorkbook(const XLWorkbook& other) = default;

        /**
         * @brief Move constructor.
         * @param other The XLWorkbook to be moved.
         */
        XLWorkbook(XLWorkbook&& other) = default;

        /**
         * @brief Destructor
         */
        ~XLWorkbook();

        /**
         * @brief Copy assignment operator.
         */
        XLWorkbook& operator=(const XLWorkbook& other) = default;

        /**
         * @brief Move assignment operator.
         */
        XLWorkbook& operator=(XLWorkbook&& other) = default;

        /**
         * @brief Get the sheet (worksheet or chartsheet) at the given index.
         * @param index The index at which the desired sheet is located.
         * @return A pointer to an XLAbstractSheet with the sheet at the index.
         * @note The index must be 1-based (rather than 0-based) as this is the default for Excel spreadsheets.
         */
        XLSheet sheet(uint16_t index);

        /**
         * @brief Get the sheet (worksheet or chartsheet) with the given name.
         * @param sheetName The name of the desired sheet.
         * @return A pointer to an XLAbstractSheet with the sheet at the index.
         */
        XLSheet sheet(std::string_view sheetName);

        /**
         * @brief Get the worksheet with the given name.
         * @param sheetName The name of the desired worksheet.
         * @return
         */
        XLWorksheet worksheet(std::string_view sheetName);

        /**
         * @brief Get the worksheet at the given index.
         * @param index The index (1-based) at which the desired sheet is located.
         * @return
         */
        XLWorksheet worksheet(uint16_t index);

        /**
         * @brief Get the chartsheet with the given name.
         * @param sheetName The name of the desired chartsheet.
         * @return
         */
        XLChartsheet chartsheet(std::string_view sheetName);

        /**
         * @brief Get the chartsheet at the given index.
         * @param index The index (1-based) at which the desired sheet is located.
         * @return
         */
        XLChartsheet chartsheet(uint16_t index);

        /**
         * @brief Get the defined names collection.
         * @return An XLDefinedNames object.
         */
        XLDefinedNames definedNames();

        /**
         * @brief Delete sheet (worksheet or chartsheet) from the workbook.
         * @param sheetName Name of the sheet to delete.
         * @throws XLException An exception will be thrown if trying to delete the last worksheet in the workbook
         */
        void deleteSheet(std::string_view sheetName);

        /**
         * @brief Add a new worksheet.
         * @param sheetName The name of the new worksheet.
         */
        void     addWorksheet(std::string_view sheetName);
        void     addChartsheet(std::string_view sheetName);

        /**
         * @brief Clone an existing sheet.
         * @param existingName The name of the sheet to clone.
         * @param newName The name of the new cloned sheet.
         */
        void cloneSheet(std::string_view existingName, std::string_view newName);

        /**
         * @brief Move a sheet to a new index.
         * @param sheetName The name of the sheet to move.
         * @param index The index (1-based) where the sheet shall be moved to.
         */
        void setSheetIndex(std::string_view sheetName, unsigned int index);

        /**
         * @brief Get the index of a sheet.
         * @param sheetName The name of the sheet.
         * @return The 1-based index of the sheet.
         */
        unsigned int indexOfSheet(std::string_view sheetName) const;

        /**
         * @brief Get the type of a sheet.
         * @param sheetName The name of the sheet.
         * @return The type of the sheet.
         */
        XLSheetType typeOfSheet(std::string_view sheetName) const;

        /**
         * @brief Get the type of a sheet at a given index.
         * @param index The 1-based index.
         * @return The type of the sheet.
         */
        XLSheetType typeOfSheet(unsigned int index) const;

        /**
         * @brief Get the total number of sheets in the workbook.
         */
        unsigned int sheetCount() const;

        /**
         * @brief Get the total number of worksheets in the workbook.
         */
        unsigned int worksheetCount() const;

        /**
         * @brief Get the total number of chartsheets in the workbook.
         */
        unsigned int chartsheetCount() const;

        /**
         * @brief Get a list of all sheet names.
         */
        std::vector<std::string> sheetNames() const;

        /**
         * @brief Get a list of all worksheet names.
         */
        std::vector<std::string> worksheetNames() const;

        /**
         * @brief Get a list of all chartsheet names.
         */
        std::vector<std::string> chartsheetNames() const;

        /**
         * @brief Check if a sheet exists.
         */
        bool sheetExists(std::string_view sheetName) const;

        /**
         * @brief Check if a worksheet exists.
         */
        bool worksheetExists(std::string_view sheetName) const;

        /**
         * @brief Check if a chartsheet exists.
         */
        bool chartsheetExists(std::string_view sheetName) const;

        /**
         * @brief Update references to a sheet after renaming.
         */
        void updateSheetReferences(std::string_view oldName, std::string_view newName);

        /**
         * @brief Delete all named ranges.
         */
        void deleteNamedRanges();

        /**
         * @brief Update the dimension (used range) for all worksheets in the workbook.
         */
        void updateWorksheetDimensions();

        /**
         * @brief Set a flag to force full calculation upon loading the file in Excel.
         */
        void setFullCalculationOnLoad();

        /**
         * @brief Protect the workbook.
         * @param lockStructure If true, the structure of the workbook is locked.
         * @param lockWindows If true, the windows of the workbook are locked.
         * @param password The password to protect the workbook with.
         */
        void protect(bool lockStructure, bool lockWindows, std::string_view password = "");

        /**
         * @brief Unprotect the workbook.
         */
        void unprotect();

        /**
         * @brief Check if the workbook is protected.
         */
        bool isProtected() const;

        /**
         * @brief Print the XML contents of the workbook.xml.
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:    // ---------- Private Member Functions ---------- //
        /**
         * @brief Create a new internal sheet ID.
         */
        uint16_t createInternalSheetID();

        /**
         * @brief Get the relationship ID for a sheet name.
         */
        std::string sheetID(std::string_view sheetName);

        /**
         * @brief Get the name for a relationship ID.
         */
        std::string sheetName(std::string_view sheetID) const;

        /**
         * @brief Get the visibility state for a relationship ID.
         */
        std::string sheetVisibility(std::string_view sheetID) const;

        /**
         * @brief Prepare sheet metadata in workbook.xml.
         */
        void prepareSheetMetadata(std::string_view sheetName, uint16_t internalID, std::string_view sheetPath = "");

        /**
         * @brief Set the name for a relationship ID.
         */
        void setSheetName(std::string_view sheetRID, std::string_view newName);

        /**
         * @brief Set the visibility state for a relationship ID.
         */
        void setSheetVisibility(std::string_view sheetRID, std::string_view state);

        /**
         * @brief Check if a sheet is the active tab.
         */
        bool sheetIsActive(std::string_view sheetRID) const;

        /**
         * @brief Set a sheet as the active tab.
         */
        bool setSheetActive(std::string_view sheetRID);

        /**
         * @brief Helper to check if a state string means visible.
         */
        bool isVisibleState(std::string_view state) const;

        /**
         * @brief Helper to check if a sheet node is visible.
         */
        bool isVisible(XMLNode const& sheetNode) const;
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLWORKBOOK_HPP
