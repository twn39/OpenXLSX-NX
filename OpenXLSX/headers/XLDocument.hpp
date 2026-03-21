#ifndef OPENXLSX_XLDOCUMENT_HPP
#define OPENXLSX_XLDOCUMENT_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <algorithm>    // std::find_if
#include <gsl/gsl>
#include <list>
#include <map>
#include <string>
#include <string_view>
#include <unordered_map>    // O(1) shared string lookup

// ===== OpenXLSX Includes ===== //
#include "IZipArchive.hpp"
#include "OpenXLSX-Exports.hpp"
#include "XLCommandQuery.hpp"
#include "XLComments.hpp"
#include "XLContentTypes.hpp"
#include "XLDrawing.hpp"
#include "XLProperties.hpp"
#include "XLRelationships.hpp"
#include "XLSharedStrings.hpp"
#include "XLStyles.hpp"
#include "XLTables.hpp"
#include "XLWorkbook.hpp"
#include "XLXmlData.hpp"
#include "XLZipArchive.hpp"
#include "XLChart.hpp"

namespace OpenXLSX
{
    constexpr const unsigned int pugi_parse_settings = pugi::parse_default | pugi::parse_ws_pcdata;    // TBD: | pugi::parse_comments

    constexpr const bool XLForceOverwrite = true;     // readability constant for 2nd parameter of XLDocument::saveAs
    constexpr const bool XLDoNotOverwrite = false;    //  "

    /**
     * @brief The XLDocumentProperties class is an enumeration of the possible properties (metadata) that can be set
     * for a XLDocument object (and .xlsx file)
     */
    enum class XLProperty {
        Title,
        Subject,
        Creator,
        Keywords,
        Description,
        LastModifiedBy,
        LastPrinted,
        CreationDate,
        ModificationDate,
        Category,
        Application,
        DocSecurity,
        ScaleCrop,
        Manager,
        Company,
        LinksUpToDate,
        SharedDoc,
        HyperlinkBase,
        HyperlinksChanged,
        AppVersion
    };

    /**
     * @brief This class encapsulates the concept of an excel file. It is different from the XLWorkbook, in that an
     * XLDocument holds an XLWorkbook together with its metadata, as well as methods for opening,
     * closing and saving the document.\n<b><em>The XLDocument is the entrypoint for clients
     * using the RapidXLSX library.</em></b>
     */
    class OPENXLSX_EXPORT XLDocument final
    {
        //----- Friends
        friend class XLXmlFile;
        friend class XLWorkbook;
        friend class XLSheet;
        friend class XLXmlData;
        friend class XLTableCollection;

        //---------- Public Member Functions
    public:
        /**
         * @brief Default constructor. Initializes an empty document package with a default ZIP archive.
         */
        explicit XLDocument(const IZipArchive& zipArchive = XLZipArchive());

        /**
         * @brief Path-based constructor. Loads an existing document or sets the target path for a new one.
         * @param docPath Non-owning view to the file path to minimize memory overhead during initialization.
         */
        explicit XLDocument(std::string_view docPath, const IZipArchive& zipArchive = XLZipArchive());

        /**
         * @brief Legacy string constructor. Maintained to resolve ambiguity with IZipArchive's template constructor 
         * when passing string literals or std::string.
         */
        explicit XLDocument(const std::string& docPath, const IZipArchive& zipArchive = XLZipArchive());

        XLDocument(const XLDocument& other) = delete;
        XLDocument(XLDocument&& other) noexcept = default;
        ~XLDocument();
        XLDocument& operator=(const XLDocument& other) = delete;
        XLDocument& operator=(XLDocument&& other) noexcept = default;

        void showWarnings();
        void suppressWarnings();

        /**
         * @brief Open an existing .xlsx package. Uses std::string_view for zero-copy path processing.
         * @param fileName Path to the file to open.
         */
        void open(std::string_view fileName);

        /**
         * @brief Initialize a new .xlsx package from a built-in template. 
         * @param fileName Target path for the new package.
         * @param forceOverwrite Prevents accidental data loss unless explicitly requested via XLForceOverwrite.
         */
        void create(std::string_view fileName, bool forceOverwrite = XLForceOverwrite);

        /**
         * @brief Create a new .xlsx file with the given name. Legacy interface.
         * @deprecated Use void create(std::string_view, bool) for explicit overwrite control.
         */
        [[deprecated]] void create(const std::string& fileName);

        /**
         * @brief Close the document and release underlying ZIP resources, ensuring file handles are closed properly.
         */
        void close();

        /**
         * @brief Save changes to the original file path.
         */
        void save();

        /**
         * @brief Persistence to a new location.
         * @param fileName Target path.
         * @param forceOverwrite Protects existing files from being overwritten without explicit intent.
         */
        void saveAs(std::string_view fileName, bool forceOverwrite = XLForceOverwrite);

        /**
         * @brief Save the document with a new name. Legacy interface.
         * @deprecated Use void saveAs(std::string_view, bool) for explicit overwrite control.
         */
        [[deprecated]] void saveAs(const std::string& fileName);

        /**
         * @brief Access filename to uniquely identify the document in the package.
         * @return The current document's filename. [[nodiscard]] is used to prevent state-querying errors.
         */
        [[nodiscard]] std::string name() const;

        /**
         * @brief Access the full filesystem path of the current document.
         * [[nodiscard]] prevents ignoring the current path reference.
         */
        [[nodiscard]] const std::string& path() const;

        /**
         * @brief Provides access to the [Content_Types].xml file, which defines the package structure and internal part mapping.
         */
        [[nodiscard]] XLContentTypes& contentTypes();

        /**
         * @brief Provides access to custom properties (metadata) that can be set for the document.
         */
        [[nodiscard]] XLCustomProperties& customProperties();

        /**
         * @brief Atomic table ID generator to ensure workbook-wide uniqueness for Excel table structures.
         * [[nodiscard]] is used to enforce correct result consumption.
         */
        [[nodiscard]] uint32_t nextTableId() const;

        /**
         * @brief Get the underlying workbook object, containing worksheets and global data state. 
         * [[nodiscard]] prevents side-effect-free querying errors.
         */
        [[nodiscard]] XLWorkbook workbook() const;

        /**
         * @brief Fetch metadata property by type (e.g., Title, Creator).
         */
        [[nodiscard]] std::string property(XLProperty prop) const;

        /**
         * @brief Update metadata property. Uses string_view to allow updating from literals without temporary allocations.
         * @param prop The property to set.
         * @param value A non-owning view to the value string.
         */
        void setProperty(XLProperty prop, std::string_view value);

        /**
         * @brief Remove specific metadata property.
         */
        void deleteProperty(XLProperty theProperty);

        /**
         * @return True if the document archive is valid and properly loaded.
         */
        [[nodiscard]] explicit operator bool() const;

        /**
         * @return True if the document is active and editable.
         */
        [[nodiscard]] bool isOpen() const;

        /**
         * @brief Provides access to shared document styles (fonts, fills, borders, cell formats).
         */
        [[nodiscard]] XLStyles& styles();

        /**
         * @brief Component presence checks, used to avoid unnecessary parsing of absent components in the package.
         */
        [[nodiscard]] bool hasSheetRelationships(uint16_t sheetXmlNo, bool isChartsheet = false) const;
        [[nodiscard]] bool hasSheetVmlDrawing(uint16_t sheetXmlNo) const;
        [[nodiscard]] bool hasSheetComments(uint16_t sheetXmlNo) const;
        [[nodiscard]] bool hasSheetDrawing(uint16_t sheetXmlNo) const;
        [[nodiscard]] bool hasSheetTables(uint16_t sheetXmlNo) const;

        XLRelationships sheetRelationships(uint16_t sheetXmlNo, bool isChartsheet = false);
        XLDrawing       sheetDrawing(uint16_t sheetXmlNo);
        XLDrawing       createDrawing();
        XLDrawing       drawing(std::string_view path);
        XLVmlDrawing    sheetVmlDrawing(uint16_t sheetXmlNo);
        XLComments      sheetComments(uint16_t sheetXmlNo);
        XLTables        sheetTables(uint16_t sheetXmlNo);
        
        class XLChart   createChart(XLChartType type = XLChartType::Bar);

        /**
         * @brief Insert image into the archive's media folder. Uses string_view for zero-copy data transfer.
         * @param name Filename in the archive.
         * @param data Raw image data as a non-owning view.
         * @return The internal package path to the image.
         */
        std::string addImage(std::string_view name, std::string_view data);

        /**
         * @brief Fetch raw image data from the document archive (xl/media/).
         */
        [[nodiscard]] std::string getImage(std::string_view path) const;

        XLRelationships drawingRelationships(std::string_view drawingPath);

        /**
         * @brief Access the low-level ZIP archive for advanced package manipulation.
         * [[nodiscard]] is used to prevent state-querying errors.
         */
        [[nodiscard]] IZipArchive&       archive() { return m_archive; }
        [[nodiscard]] const IZipArchive& archive() const { return m_archive; }

        /**
         * @brief Access application-specific properties (metadata).
         * [[nodiscard]] is used to prevent state-querying errors.
         */
        [[nodiscard]] XLAppProperties&       appProperties() { return m_appProperties; }
        [[nodiscard]] const XLAppProperties& appProperties() const { return m_appProperties; }

        /**
         * @brief Access core document metadata properties.
         * [[nodiscard]] is used to prevent state-querying errors.
         */
        [[nodiscard]] XLProperties&       coreProperties() { return m_coreProperties; }
        [[nodiscard]] const XLProperties& coreProperties() const { return m_coreProperties; }

        /**
         * @brief Ensure worksheet names adhere to Excel's naming rules.
         */
        bool validateSheetName(std::string_view sheetName, bool throwOnInvalid = false);

        bool execCommand(const XLCommand& command);

        /**
         * @brief Execute a state-querying operation without modifying the document.
         * [[nodiscard]] is used to prevent state-querying errors.
         */
        [[nodiscard]] XLQuery execQuery(const XLQuery& query) const;
        XLQuery               execQuery(const XLQuery& query);

        void setSavingDeclaration(XLXmlSavingDeclaration const& savingDeclaration);

        /**
         * @brief Access the workbook's central Shared String Table (SST). 
         * The SST reduces file size by deduplicating text across the entire package.
         */
        [[nodiscard]] const XLSharedStrings& sharedStrings() const { return m_sharedStrings; }

        /**
         * @brief Prune unused strings from the SST and reindex referencing cells. 
         * Essential for minimizing file size and memory footprint in long-lived or large documents.
         */
        void cleanupSharedStrings();

    protected:
        /**
         * @brief Fetch raw XML content for a specific package path. 
         * Used internally for lazily loading components on demand.
         */
        [[nodiscard]] std::string extractXmlFromArchive(std::string_view path);

        /**
         * @brief Provide access to managed XML data nodes, enabling centralized XML state management.
         * doNotThrow allows safe querying when a component's presence is uncertain.
         */
        [[nodiscard]] XLXmlData*       getXmlData(std::string_view path, bool doNotThrow = false);
        [[nodiscard]] const XLXmlData* getXmlData(std::string_view path, bool doNotThrow = false) const;
        [[nodiscard]] bool             hasXmlData(std::string_view path) const;


    private:
        bool        m_suppressWarnings{true};
        std::string m_filePath{};

        XLXmlSavingDeclaration m_xmlSavingDeclaration;

        struct SharedFormula
        {
            std::string formula;
            uint32_t    baseRow;
            uint16_t    baseCol;
        };

        mutable std::list<XLXmlData>                     m_data{};
        mutable std::deque<std::string>                  m_sharedStringCache{};
        mutable std::unordered_map<std::string, int32_t> m_sharedStringIndex{};
        mutable XLSharedStrings                          m_sharedStrings{};
        mutable std::map<const XMLDocument*, std::unordered_map<uint32_t, SharedFormula>>
            m_sharedFormulas{};
        std::map<std::string, std::string> m_unhandledEntries{};

        XLRelationships    m_docRelationships{};
        XLRelationships    m_wbkRelationships{};
        XLContentTypes     m_contentTypes{};
        XLAppProperties    m_appProperties{};
        XLProperties       m_coreProperties{};
        XLCustomProperties m_customProperties{};
        XLStyles           m_styles{};
        XLWorkbook         m_workbook{};
        IZipArchive        m_archive{};
    };

    [[nodiscard]] OPENXLSX_EXPORT std::string BinaryAsHexString(gsl::span<const std::byte> data);
    [[nodiscard]] OPENXLSX_EXPORT uint16_t    ExcelPasswordHash(std::string_view password);
    [[nodiscard]] OPENXLSX_EXPORT std::string ExcelPasswordHashAsString(std::string_view password);

    [[nodiscard]] std::string getPathARelativeToPathB(std::string_view pathA, std::string_view pathB);
    [[nodiscard]] std::string eliminateDotAndDotDotFromPath(std::string_view path);

}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLDOCUMENT_HPP
