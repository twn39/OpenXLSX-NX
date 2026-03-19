#ifndef OPENXLSX_XLSHEET_HPP
#define OPENXLSX_XLSHEET_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(push)
#    pragma warning(disable : 4251)
#    pragma warning(disable : 4275)
#endif    // _MSC_VER

// ===== External Includes ===== //
#include <ostream>
#include <variant>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLSheetBase.hpp"
#include "XLConditionalFormatting.hpp"
#include "XLWorksheet.hpp"
#include "XLChartsheet.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    /**
     * @brief The XLAbstractSheet is a generalized sheet class, which functions as superclass for specialized classes,
     * such as XLWorksheet. It implements functionality common to all sheet types. This is a pure abstract class,
     * so it cannot be instantiated.
     */
    class OPENXLSX_EXPORT XLSheet final : public XLXmlFile
    {
    public:
        /**
         * @brief The constructor. There are no default constructor, so all parameters must be provided for
         * constructing an XLAbstractSheet object. Since this is a pure abstract class, instantiation is only
         * possible via one of the derived classes.
         * @param xmlData
         */
        explicit XLSheet(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLSheet(const XLSheet& other) = default;

        /**
         * @brief
         * @param other
         */
        XLSheet(XLSheet&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLSheet() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLSheet& operator=(const XLSheet& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSheet& operator=(XLSheet&& other) noexcept = default;

        /**
         * @brief Method for getting the current visibility state of the sheet.
         * @return An XLSheetState enum object, with the current sheet state.
         */
        XLSheetState visibility() const;

        /**
         * @brief Method for setting the state of the sheet.
         * @param state An XLSheetState enum object with the new state.
         * @bug For some reason, this method doesn't work. The data is written correctly to the xml file, but the sheet
         * is not hidden when opening the file in Excel.
         */
        void setVisibility(XLSheetState state);

        /**
         * @brief
         * @return
         */
        XLColor color() const;

        /**
         * @brief
         * @param color
         */
        void setColor(const XLColor& color);

        /**
         * @brief Method for getting the index of the sheet.
         * @return An int with the index of the sheet.
         */
        uint16_t index() const;

        /**
         * @brief Method for setting the index of the sheet. This effectively moves the sheet to a different position.
         */
        void setIndex(uint16_t index);

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        std::string name() const;

        /**
         * @brief Method for renaming the sheet.
         * @param name A std::string with the new name.
         */
        void setName(const std::string& name);

        /**
         * @brief Determine whether the sheet is selected
         * @return
         */
        bool isSelected() const;

        /**
         * @brief
         * @param selected
         */
        void setSelected(bool selected);

        /**
         * @brief
         * @return
         */
        bool isActive() const;

        /**
         * @brief
         */
        bool setActive();

        /**
         * @brief Method to get the type of the sheet.
         * @return An XLSheetType enum object with the sheet type.
         */
        template<typename SheetType,
                 typename = std::enable_if_t<std::is_same_v<SheetType, XLWorksheet> or std::is_same_v<SheetType, XLChartsheet>>>
        bool isType() const
        { return std::holds_alternative<SheetType>(m_sheet); }

        /**
         * @brief Method for cloning the sheet.
         * @param newName A std::string with the name of the clone
         * @return A pointer to the cloned object.
         * @note This is a pure abstract method. I.e. it is implemented in subclasses.
         */
        void clone(const std::string& newName);

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T, typename = std::enable_if_t<std::is_same_v<T, XLWorksheet> or std::is_same_v<T, XLChartsheet>>>
        T get() const
        {
            try {
                if constexpr (std::is_same<T, XLWorksheet>::value)
                    return std::get<XLWorksheet>(m_sheet);

                else if constexpr (std::is_same<T, XLChartsheet>::value)
                    return std::get<XLChartsheet>(m_sheet);
            }

            catch (const std::bad_variant_access&) {
                throw XLSheetError("XLSheet object does not contain the requested sheet type.");
            }
        }

        /**
         * @brief
         * @return
         */
        operator XLWorksheet() const;    // NOLINT

        /**
         * @brief
         * @return
         */
        operator XLChartsheet() const;    // NOLINT

        /**
         * @brief Print the XML contents of the XLSheet using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:                                             // ---------- Private Member Variables ---------- //
        std::variant<XLWorksheet, XLChartsheet> m_sheet; /**<  */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#    pragma warning(pop)
#endif    // _MSC_VER

#endif    // OPENXLSX_XLSHEET_HPP
