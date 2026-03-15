#ifndef OPENXLSX_XLPROPERTIES_HPP
#define OPENXLSX_XLPROPERTIES_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <string>
#include <string_view>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"
#include "XLDateTime.hpp"

namespace OpenXLSX
{
    /**
     * @brief The XLProperties class represents the core properties of the document (docProps/core.xml).
     */
    class OPENXLSX_EXPORT XLProperties : public XLXmlFile
    {
    private:
        /**
         * @brief constructor helper function: create core.xml content from template
         */
        void createFromTemplate();

    public:
        /**
         * @brief Default constructor.
         */
        XLProperties() = default;

        /**
         * @brief Constructor that binds to existing XML data.
         * @param xmlData Pointer to the XML data.
         */
        explicit XLProperties(XLXmlData* xmlData);

        /**
         * @brief Copy constructor.
         */
        XLProperties(const XLProperties& other) = default;

        /**
         * @brief Move constructor.
         */
        XLProperties(XLProperties&& other) noexcept = default;

        /**
         * @brief Destructor.
         */
        ~XLProperties();

        /**
         * @brief Copy assignment operator.
         */
        XLProperties& operator=(const XLProperties& other) = default;

        /**
         * @brief Move assignment operator.
         */
        XLProperties& operator=(XLProperties&& other) noexcept = default;

        /**
         * @brief Set a property by name.
         * @param name The name of the property (including namespace prefix, e.g., "dc:title").
         * @param value The value to set.
         */
        void setProperty(std::string_view name, std::string_view value);

        /**
         * @brief Set a numeric property.
         * @param name The name of the property.
         * @param value The integer value.
         */
        void setProperty(std::string_view name, int value);

        /**
         * @brief Set a numeric property.
         * @param name The name of the property.
         * @param value The double value.
         */
        void setProperty(std::string_view name, double value);

        /**
         * @brief Get a property value as a string.
         * @param name The name of the property.
         * @return The property value.
         */
        std::string property(std::string_view name) const;

        /**
         * @brief Delete a property.
         * @param name The name of the property.
         */
        void deleteProperty(std::string_view name);

        // Standard Core Properties
        std::string title() const;
        void setTitle(std::string_view title);

        std::string subject() const;
        void setSubject(std::string_view subject);

        std::string creator() const;
        void setCreator(std::string_view creator);

        std::string keywords() const;
        void setKeywords(std::string_view keywords);

        std::string description() const;
        void setDescription(std::string_view description);

        std::string lastModifiedBy() const;
        void setLastModifiedBy(std::string_view author);

        std::string revision() const;
        void setRevision(std::string_view revision);

        std::string lastPrinted() const;
        void setLastPrinted(std::string_view date);

        XLDateTime created() const;
        void setCreated(const XLDateTime& date);

        XLDateTime modified() const;
        void setModified(const XLDateTime& date);

        std::string category() const;
        void setCategory(std::string_view category);

        std::string contentStatus() const;
        void setContentStatus(std::string_view status);
    };

    /**
     * @brief The XLAppProperties class represents the extended/application properties (docProps/app.xml).
     */
    class OPENXLSX_EXPORT XLAppProperties : public XLXmlFile
    {
    private:
        /**
         * @brief constructor helper function: create app.xml content from template
         * @param workbookXml The workbook XML document.
         */
        void createFromTemplate(XMLDocument const& workbookXml);

    public:
        /**
         * @brief Default constructor.
         */
        XLAppProperties() = default;

        /**
         * @brief Constructor for creating from existing data and workbook.
         * @param xmlData Pointer to the XML data.
         * @param workbookXml The workbook XML document.
         */
        explicit XLAppProperties(XLXmlData* xmlData, XMLDocument const& workbookXml);

        /**
         * @brief Constructor for existing XML data.
         * @param xmlData Pointer to the XML data.
         */
        explicit XLAppProperties(XLXmlData* xmlData);

        /**
         * @brief Copy constructor.
         */
        XLAppProperties(const XLAppProperties& other) = default;

        /**
         * @brief Move constructor.
         */
        XLAppProperties(XLAppProperties&& other) noexcept = default;

        /**
         * @brief Destructor.
         */
        ~XLAppProperties();

        /**
         * @brief Copy assignment operator.
         */
        XLAppProperties& operator=(const XLAppProperties& other) = default;

        /**
         * @brief Move assignment operator.
         */
        XLAppProperties& operator=(XLAppProperties&& other) noexcept = default;

        /**
         * @brief Increment or decrement the worksheet count in HeadingPairs and TitlesOfParts.
         * @param increment The change in count.
         */
        void incrementSheetCount(int16_t increment);

        /**
         * @brief Align the sheet names in TitlesOfParts with the actual workbook sheet names.
         * @param workbookSheetNames Vector of sheet names.
         */
        void alignWorksheets(const std::vector<std::string>& workbookSheetNames);

        /**
         * @brief Add a sheet name to TitlesOfParts.
         * @param title The sheet name.
         */
        void addSheetName(std::string_view title);

        /**
         * @brief Delete a sheet name from TitlesOfParts.
         * @param title The sheet name.
         */
        void deleteSheetName(std::string_view title);

        /**
         * @brief Update a sheet name in TitlesOfParts.
         * @param oldTitle The old name.
         * @param newTitle The new name.
         */
        void setSheetName(std::string_view oldTitle, std::string_view newTitle);

        /**
         * @brief Add a HeadingPair (e.g., "Worksheets").
         * @param name The name of the pair.
         * @param value The value.
         */
        void addHeadingPair(std::string_view name, int value);

        /**
         * @brief Delete a HeadingPair.
         * @param name The name of the pair.
         */
        void deleteHeadingPair(std::string_view name);

        /**
         * @brief Update a HeadingPair value.
         * @param name The name of the pair.
         * @param newValue The new value.
         */
        void setHeadingPair(std::string_view name, int newValue);

        /**
         * @brief Set a generic property in app.xml.
         * @param name The property name.
         * @param value The value.
         */
        void setProperty(std::string_view name, std::string_view value);

        /**
         * @brief Get a property value from app.xml.
         * @param name The property name.
         * @return The property value.
         */
        std::string property(std::string_view name) const;

        /**
         * @brief Delete a property from app.xml.
         * @param name The property name.
         */
        void deleteProperty(std::string_view name);

        /**
         * @brief Append a sheet name.
         * @param sheetName The name to append.
         */
        void appendSheetName(std::string_view sheetName);

        /**
         * @brief Prepend a sheet name.
         * @param sheetName The name to prepend.
         */
        void prependSheetName(std::string_view sheetName);

        /**
         * @brief Insert a sheet name at a specific index.
         * @param sheetName The name to insert.
         * @param index The 1-based index.
         */
        void insertSheetName(std::string_view sheetName, unsigned int index);
    };

    /**
     * @brief The XLCustomProperties class represents custom user-defined properties (docProps/custom.xml).
     */
    class OPENXLSX_EXPORT XLCustomProperties : public XLXmlFile
    {
    private:
        /**
         * @brief Initialize from template.
         */
        void createFromTemplate();

    public:
        /**
         * @brief Default constructor.
         */
        XLCustomProperties() = default;

        /**
         * @brief Constructor for existing XML data.
         * @param xmlData Pointer to the XML data.
         */
        explicit XLCustomProperties(XLXmlData* xmlData);

        /**
         * @brief Copy constructor.
         */
        XLCustomProperties(const XLCustomProperties& other) = default;

        /**
         * @brief Move constructor.
         */
        XLCustomProperties(XLCustomProperties&& other) noexcept = default;

        /**
         * @brief Destructor.
         */
        ~XLCustomProperties();

        /**
         * @brief Copy assignment operator.
         */
        XLCustomProperties& operator=(const XLCustomProperties& other) = default;

        /**
         * @brief Move assignment operator.
         */
        XLCustomProperties& operator=(XLCustomProperties&& other) noexcept = default;

        /**
         * @brief Set a string custom property (vt:lpwstr).
         * @param name The property name.
         * @param value The value.
         */
        void setProperty(std::string_view name, std::string_view value);

        /**
         * @brief Overload for string literals.
         * @param name The property name.
         * @param value The literal value.
         */
        void setProperty(std::string_view name, const char* value);

        /**
         * @brief Set an integer custom property (vt:i4).
         * @param name The property name.
         * @param value The value.
         */
        void setProperty(std::string_view name, int value);

        /**
         * @brief Set a double custom property (vt:r8).
         * @param name The property name.
         * @param value The value.
         */
        void setProperty(std::string_view name, double value);

        /**
         * @brief Set a boolean custom property (vt:bool).
         * @param name The property name.
         * @param value The value.
         */
        void setProperty(std::string_view name, bool value);

        /**
         * @brief Set a date/time custom property (vt:filetime).
         * @param name The property name.
         * @param value The XLDateTime value.
         */
        void setProperty(std::string_view name, const XLDateTime& value);

        /**
         * @brief Get a custom property value as a string.
         * @param name The property name.
         * @return The value as string.
         */
        std::string property(std::string_view name) const;

        /**
         * @brief Delete a custom property.
         * @param name The property name.
         */
        void deleteProperty(std::string_view name);
    };

}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLPROPERTIES_HPP
