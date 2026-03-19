#ifndef OPENXLSX_XLSHEETBASE_HPP
#define OPENXLSX_XLSHEETBASE_HPP

#include <cstdint>
#include <string>
#include <type_traits>
#include "OpenXLSX-Exports.hpp"
#include "XLColor.hpp"
#include "XLCommandQuery.hpp"
#include "XLDocument.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    class XLWorksheet;
    class XLChartsheet;

    constexpr const bool XLEmptyHiddenCells = true;    // readability constant for XLWorksheet::mergeCells parameter emptyHiddenCells

    /**
     * @brief The XLSheetState is an enumeration of the possible (visibility) states, e.g. Visible or Hidden.
     */
    enum class XLSheetState : uint8_t { Visible, Hidden, VeryHidden };

    /**
     * @brief The XLPaneState is an enumeration of the possible states of a pane, e.g. Frozen or Split.
     */
    enum class XLPaneState : uint8_t { Split, Frozen, FrozenSplit };

    /**
     * @brief The XLPane is an enumeration of the possible pane identifiers.
     */
    enum class XLPane : uint8_t { BottomRight, TopRight, BottomLeft, TopLeft };

    OPENXLSX_EXPORT std::string XLPaneStateToString(XLPaneState state);
    OPENXLSX_EXPORT XLPaneState XLPaneStateFromString(std::string const& stateString);
    OPENXLSX_EXPORT std::string XLPaneToString(XLPane pane);
    OPENXLSX_EXPORT XLPane      XLPaneFromString(std::string const& paneString);

    OPENXLSX_EXPORT void setTabColor(const XMLDocument& xmlDocument, const XLColor& color);
    OPENXLSX_EXPORT void setTabSelected(const XMLDocument& xmlDocument, bool selected);
    OPENXLSX_EXPORT bool tabIsSelected(const XMLDocument& xmlDocument);

    /**
     * @brief The XLSheetBase class is the base class for the XLWorksheet and XLChartsheet classes. However,
     * it is not a base class in the traditional sense. Rather, it provides common functionality that is
     * inherited via the CRTP (Curiously Recurring Template Pattern) pattern.
     * @tparam T Type that will inherit functionality. Restricted to types XLWorksheet and XLChartsheet.
     */
    template<typename T, typename = std::enable_if_t<std::is_same_v<T, XLWorksheet> or std::is_same_v<T, XLChartsheet>>>
    class OPENXLSX_EXPORT XLSheetBase : public XLXmlFile
    {
    public:
        /**
         * @brief Constructor
         */
        XLSheetBase() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor. There are no default constructor, so all parameters must be provided for
         * constructing an XLAbstractSheet object. Since this is a pure abstract class, instantiation is only
         * possible via one of the derived classes.
         * @param xmlData
         */
        explicit XLSheetBase(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLSheetBase(const XLSheetBase& other) = default;

        /**
         * @brief
         * @param other
         */
        XLSheetBase(XLSheetBase&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLSheetBase() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLSheetBase& operator=(const XLSheetBase&) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSheetBase& operator=(XLSheetBase&& other) noexcept = default;

        /**
         * @brief
         * @return
         */
        XLSheetState visibility() const
        {
            XLQuery query(XLQueryType::QuerySheetVisibility);
            query.setParam("sheetID", relationshipID());
            const auto state  = parentDoc().execQuery(query).template result<std::string>();
            auto       result = XLSheetState::Visible;

            if (state == "visible" or state.empty()) { result = XLSheetState::Visible; }
            else if (state == "hidden") {
                result = XLSheetState::Hidden;
            }
            else if (state == "veryHidden") {
                result = XLSheetState::VeryHidden;
            }

            return result;
        }

        /**
         * @brief
         * @param state
         */
        void setVisibility(XLSheetState state)
        {
            auto stateString = std::string();
            switch (state) {
                case XLSheetState::Visible:
                    stateString = "visible";
                    break;

                case XLSheetState::Hidden:
                    stateString = "hidden";
                    break;

                case XLSheetState::VeryHidden:
                    stateString = "veryHidden";
                    break;
            }

            parentDoc().execCommand(XLCommand(XLCommandType::SetSheetVisibility)
                                        .setParam("sheetID", relationshipID())
                                        .setParam("sheetVisibility", stateString));
        }

        /**
         * @brief
         * @return
         * @todo To be implemented.
         */
        XLColor color() const { return static_cast<const T&>(*this).getColor_impl(); }

        /**
         * @brief
         * @param color
         */
        void setColor(const XLColor& color) { static_cast<T&>(*this).setColor_impl(color); }

        /**
         * @brief look up sheet name via relationshipID, then attempt to match that to a sheet in XLWorkbook::sheet(uint16_t index), looping
         * over index
         * @return the index by which the sheet could be accessed from XLWorkbook::sheet
         */
        uint16_t index() const
        {
            XLQuery query(XLQueryType::QuerySheetIndex);
            query.setParam("sheetID", relationshipID());
            return static_cast<uint16_t>(std::stoi(parentDoc().execQuery(query).template result<std::string>()));
        }

        /**
         * @brief
         * @param index
         */
        void setIndex(uint16_t index)
        {
            parentDoc().execCommand(
                XLCommand(XLCommandType::SetSheetIndex).setParam("sheetID", relationshipID()).setParam("sheetIndex", index));
        }

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        std::string name() const
        {
            XLQuery query(XLQueryType::QuerySheetName);
            query.setParam("sheetID", relationshipID());
            return parentDoc().execQuery(query).template result<std::string>();
        }

        /**
         * @brief Method for renaming the sheet.
         * @param sheetName A std::string with the new name.
         */
        void setName(const std::string& sheetName)
        {
            parentDoc().execCommand(XLCommand(XLCommandType::SetSheetName)
                                        .setParam("sheetID", relationshipID())
                                        .setParam("sheetName", name())
                                        .setParam("newName", sheetName));
        }

        /**
         * @brief
         * @return
         */
        bool isSelected() const { return static_cast<const T&>(*this).isSelected_impl(); }

        /**
         * @brief
         * @param selected
         */
        void setSelected(bool selected) { static_cast<T&>(*this).setSelected_impl(selected); }

        /**
         * @brief
         * @return
         */
        bool isActive() const { return static_cast<const T&>(*this).isActive_impl(); }

        /**
         * @brief
         */
        bool setActive() { return static_cast<T&>(*this).setActive_impl(); }

        /**
         * @brief Method for cloning the sheet.
         * @param newName A std::string with the name of the clone
         * @return A pointer to the cloned object.
         * @note This is a pure abstract method. I.e. it is implemented in subclasses.
         */
        void clone(const std::string& newName)
        {
            parentDoc().execCommand(
                XLCommand(XLCommandType::CloneSheet).setParam("sheetID", relationshipID()).setParam("cloneName", newName));
        }

    private:    // ---------- Private Member Variables ---------- //
    };
}

#endif
