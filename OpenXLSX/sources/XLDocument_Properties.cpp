#include <string>
#include <string_view>
#include <charconv>
#include <fmt/format.h>

#include "XLDocument.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details Provides read access to standard OOXML document metadata (core and extended properties) such as creator, application version, or creation date.
 */
std::string XLDocument::property(XLProperty prop) const
{
    switch (prop) {
        case XLProperty::Application:
            return m_appProperties.property("Application");
        case XLProperty::AppVersion:
            return m_appProperties.property("AppVersion");
        case XLProperty::Category:
            return m_coreProperties.category();
        case XLProperty::Company:
            return m_appProperties.property("Company");
        case XLProperty::CreationDate:
            return m_coreProperties.property("dcterms:created");
        case XLProperty::Creator:
            return m_coreProperties.creator();
        case XLProperty::Description:
            return m_coreProperties.description();
        case XLProperty::DocSecurity:
            return m_appProperties.property("DocSecurity");
        case XLProperty::HyperlinkBase:
            return m_appProperties.property("HyperlinkBase");
        case XLProperty::HyperlinksChanged:
            return m_appProperties.property("HyperlinksChanged");
        case XLProperty::Keywords:
            return m_coreProperties.keywords();
        case XLProperty::LastModifiedBy:
            return m_coreProperties.lastModifiedBy();
        case XLProperty::LastPrinted:
            return m_coreProperties.lastPrinted();
        case XLProperty::LinksUpToDate:
            return m_appProperties.property("LinksUpToDate");
        case XLProperty::Manager:
            return m_appProperties.property("Manager");
        case XLProperty::ModificationDate:
            return m_coreProperties.property("dcterms:modified");
        case XLProperty::ScaleCrop:
            return m_appProperties.property("ScaleCrop");
        case XLProperty::SharedDoc:
            return m_appProperties.property("SharedDoc");
        case XLProperty::Subject:
            return m_coreProperties.subject();
        case XLProperty::Title:
            return m_coreProperties.title();
        default:
            return "";    // To silence compiler warning.
    }
}

/**
 * @brief extract an integer major version v1 and minor version v2 from a string
 * @details Extracts integer major and minor versions from a string, ensuring the formatting strictly complies with the OOXML AppVersion schema requirements.
 * [0-9]{1,2}\.[0-9]{1,4}  (Example: 14.123)
 * @param versionString the string to process
 * @param majorVersion by reference: store the major version here
 * @param minorVersion by reference: store the minor version here
 * @return true if string adheres to format & version numbers could be extracted
 * @return false in case of failure
 */
bool getAppVersion(std::string_view versionString, int& majorVersion, int& minorVersion)
{
    // ===== const expressions for hardcoded version limits
    constexpr int minMajorV = 0, maxMajorV = 99;      // allowed value range for major version number
    constexpr int minMinorV = 0, maxMinorV = 9999;    //          "          for minor   "

    const size_t begin  = versionString.find_first_not_of(" \t");
    const size_t dotPos = versionString.find_first_of('.');

    // early failure if string is only blanks or does not contain a dot
    if (begin == std::string_view::npos or dotPos == std::string_view::npos) return false;

    const size_t end = versionString.find_last_not_of(" \t");
    if (begin != std::string_view::npos and dotPos != std::string_view::npos) {
        std::string_view strMajorVersion = versionString.substr(begin, dotPos - begin);
        std::string_view strMinorVersion = versionString.substr(dotPos + 1, end - dotPos);
        
        auto resMajor = std::from_chars(strMajorVersion.data(), strMajorVersion.data() + strMajorVersion.size(), majorVersion);
        if (resMajor.ec != std::errc() || resMajor.ptr != strMajorVersion.data() + strMajorVersion.size()) return false;
        
        auto resMinor = std::from_chars(strMinorVersion.data(), strMinorVersion.data() + strMinorVersion.size(), minorVersion);
        if (resMinor.ec != std::errc() || resMinor.ptr != strMinorVersion.data() + strMinorVersion.size()) return false;
    }
    if (majorVersion < minMajorV or majorVersion > maxMajorV or minorVersion < minMinorV ||
        minorVersion > maxMinorV)    // final range check
        return false;

    return true;
}

/**
 * @details Mutates document metadata, explicitly validating format-specific constraints (e.g., boolean strings or version formatting 'XX.XXXX') to prevent schema validation failures in Excel.
 */
void XLDocument::setProperty(XLProperty prop, std::string_view value)
{
    const std::string valStr(value);
    switch (prop) {
        case XLProperty::Application:
            m_appProperties.setProperty("Application", valStr);
            break;
        case XLProperty::AppVersion: {
            int minorVersion, majorVersion;
            // ===== Check for the format "XX.XXXX", with X being a number.
            if (!getAppVersion(value, majorVersion, minorVersion)) throw XLPropertyError("Invalid property value: " + valStr);
            m_appProperties.setProperty("AppVersion", fmt::format("{}.{}", majorVersion, minorVersion));
            break;
        }

        case XLProperty::Category:
            m_coreProperties.setCategory(valStr);
            break;
        case XLProperty::Company:
            m_appProperties.setProperty("Company", valStr);
            break;
        case XLProperty::CreationDate:
            m_coreProperties.setProperty("dcterms:created", valStr);
            break;
        case XLProperty::Creator:
            m_coreProperties.setCreator(valStr);
            break;
        case XLProperty::Description:
            m_coreProperties.setDescription(valStr);
            break;
        case XLProperty::DocSecurity:
            if (value == "0" or value == "1" or value == "2" or value == "4" or value == "8")
                m_appProperties.setProperty("DocSecurity", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::HyperlinkBase:
            m_appProperties.setProperty("HyperlinkBase", valStr);
            break;
        case XLProperty::HyperlinksChanged:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("HyperlinksChanged", valStr);
            else
                throw XLPropertyError("Invalid property value");

            break;

        case XLProperty::Keywords:
            m_coreProperties.setKeywords(valStr);
            break;
        case XLProperty::LastModifiedBy:
            m_coreProperties.setLastModifiedBy(valStr);
            break;
        case XLProperty::LastPrinted:
            m_coreProperties.setLastPrinted(valStr);
            break;
        case XLProperty::LinksUpToDate:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("LinksUpToDate", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::Manager:
            m_appProperties.setProperty("Manager", valStr);
            break;
        case XLProperty::ModificationDate:
            m_coreProperties.setProperty("dcterms:modified", valStr);
            break;
        case XLProperty::ScaleCrop:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("ScaleCrop", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::SharedDoc:
            if (value == "true" or value == "false")
                m_appProperties.setProperty("SharedDoc", valStr);
            else
                throw XLPropertyError("Invalid property value");
            break;

        case XLProperty::Subject:
            m_coreProperties.setSubject(valStr);
            break;
        case XLProperty::Title:
            m_coreProperties.setTitle(valStr);
            break;
    }
}

/**
 * @details Clears a specific metadata property to clean up unneeded, obsolete, or sensitive document traces.
 */
void XLDocument::deleteProperty(XLProperty theProperty) { setProperty(theProperty, ""); }
