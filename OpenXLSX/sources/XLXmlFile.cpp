// ===== External Includes ===== //
#include <gsl/gsl>
#include <pugixml.hpp>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLPackagePartFactory.hpp"
#include "XLPackageServices.hpp"
#include "XLSharedStringTable.hpp"
#include "XLXmlFile.hpp"

using namespace OpenXLSX;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. If the XML data is provided by a string, any file with
 * the same path in the .zip file will be overwritten upon saving of the document. If no xmlData is provided,
 * the data will be read from the .zip file, using the given path.
 */
XLXmlFile::XLXmlFile(XLXmlData* xmlData) : m_xmlData(xmlData) {}

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 * When envoking the load_string method in PugiXML, the flag 'parse_ws_pcdata' is passed along with the default flags.
 * This will enable parsing of whitespace characters. If not set, Excel cells with only spaces will be returned as
 * empty strings, which is not what we want. The downside is that whitespace characters such as \\n and \\t in the
 * input xml file may mess up the parsing.
 */
void XLXmlFile::setXmlData(std::string_view xmlData)
{
    Expects(m_xmlData != nullptr);
    m_xmlData->setRawData(std::string(xmlData));
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
std::string XLXmlFile::xmlData(XLXmlSavingDeclaration savingDeclaration) const
{
    Expects(m_xmlData != nullptr);
    return m_xmlData->getRawData(savingDeclaration);
}

const XLDocument& XLXmlFile::parentDoc() const
{
    Expects(m_xmlData != nullptr);
    return *m_xmlData->getParentDoc();
}

XLDocument& XLXmlFile::parentDoc()
{
    Expects(m_xmlData != nullptr);
    return *m_xmlData->getParentDoc();
}

XLPackageServices& XLXmlFile::package()
{
    // XLDocument implements XLPackageServices; prefer this over parentDoc() for package I/O.
    return parentDoc();
}

const XLPackageServices& XLXmlFile::package() const
{
    return parentDoc();
}

XLPackagePartFactory& XLXmlFile::partFactory()
{
    // XLDocument implements XLPackagePartFactory; prefer for createChart / drawings / slicers.
    return parentDoc();
}

const XLPackagePartFactory& XLXmlFile::partFactory() const
{
    return parentDoc();
}

const XLSharedStringTable& XLXmlFile::sharedStringTable() const
{
    // XLDocument implements XLSharedStringTable (read-only SST lookups).
    return parentDoc();
}

const XLSharedStrings& XLXmlFile::sharedStrings() const
{
    // Full SST (write-capable via const methods on XLSharedStrings). Prefer over parentDoc().sharedStrings().
    return parentDoc().sharedStrings();
}

std::string XLXmlFile::relationshipID() const
{
    Expects(m_xmlData != nullptr);
    return m_xmlData->getXmlID();
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource.
 */
XMLDocument& XLXmlFile::xmlDocument()
{
    Expects(m_xmlData != nullptr);
    return *m_xmlData->getXmlDocument();
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource as const.
 */
const XMLDocument& XLXmlFile::xmlDocument() const
{
    Expects(m_xmlData != nullptr);
    return *m_xmlData->getXmlDocument();
}

/**
 * @details provide access to the underlying XLXmlData::getXmlPath() function
 */
std::string XLXmlFile::getXmlPath() const { return m_xmlData == nullptr ? "" : m_xmlData->getXmlPath(); }
