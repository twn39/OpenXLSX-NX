// ===== External Includes ===== //
#include <pugixml.hpp>
#include <string_view>
#include <vector>
#include <algorithm>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLTables.hpp"
#include "XLUtilities.hpp"
#include "XLSheet.hpp"
#include "XLCellRange.hpp"

namespace OpenXLSX
{
    /**
     * @brief The TableNodeOrder vector defines the order of nodes in the table XML file.
     */
    const std::vector<std::string_view> TableNodeOrder = {
        "autoFilter",
        "sortState",
        "tableColumns",
        "tableStyleInfo",
        "extLst"
    };

    XLTableCollection::XLTableCollection(XMLNode node, XLWorksheet* worksheet) : m_sheetNode(node), m_worksheet(worksheet) {}

    void XLTableCollection::load() const
    {
        if (m_loaded || m_worksheet == nullptr) return;
        m_tables.clear();
        m_loaded = true;

        auto tableParts = m_sheetNode.child("tableParts");
        if (tableParts.empty()) return;

        for (auto& tablePart : tableParts.children("tablePart")) {
            std::string rId    = tablePart.attribute("r:id").value();
            std::string target = m_worksheet->relationships().relationshipById(rId).target();
            std::string worksheetPath = m_worksheet->getXmlPath();
            std::string absolutePath  = eliminateDotAndDotDotFromPath(!target.empty() && target.front() == '/' ? target : worksheetPath.substr(0, worksheetPath.find_last_of('/') + 1) + target);
            if (!absolutePath.empty() && absolutePath.front() == '/') absolutePath.erase(0, 1);
            XLXmlData* xmlData = m_worksheet->parentDoc().getXmlData(absolutePath, true);
            if (!xmlData && m_worksheet->parentDoc().archive().hasEntry(absolutePath)) {
                m_worksheet->parentDoc().contentTypes().addOverride("/" + absolutePath, XLContentType::Table);
                xmlData = &m_worksheet->parentDoc().m_data.emplace_back(&m_worksheet->parentDoc(), absolutePath, "", XLContentType::Table);
            }
            if (xmlData) m_tables.emplace_back(xmlData);
        }
    }

    size_t XLTableCollection::count() const { load(); return m_tables.size(); }
    XLTable XLTableCollection::operator[](size_t index) const { load(); return m_tables.at(index); }
    XLTable XLTableCollection::table(std::string_view name) const { load(); auto it = std::find_if(m_tables.begin(), m_tables.end(), [&](const XLTable& t) { return t.name() == name; }); if (it == m_tables.end()) throw XLException("Table not found"); return *it; }

    XLTable XLTableCollection::add(std::string_view name, std::string_view range)
    {
        uint32_t tableId = m_worksheet->parentDoc().nextTableId();
        std::string tablesFilename = "xl/tables/table" + std::to_string(tableId) + ".xml";
        std::string tablesRelativePath = getPathARelativeToPathB(tablesFilename, m_worksheet->getXmlPath());
        
        auto rel = m_worksheet->relationships().addRelationship(XLRelationshipType::Table, tablesRelativePath);
        
        auto tableParts = m_sheetNode.child("tableParts");
        if (tableParts.empty()) tableParts = appendAndGetNode(m_sheetNode, "tableParts", m_worksheet->m_nodeOrder);
        
        uint32_t count = tableParts.attribute("count").as_uint(0) + 1;
        appendAndSetAttribute(tableParts, "count", std::to_string(count));
        tableParts.append_child("tablePart").append_attribute("r:id").set_value(rel.id().c_str());

        std::string refStr(range);
        uint32_t colCount = 1;
        XLCellReference start("A1");
        auto colonPos = refStr.find(':');
        if (colonPos != std::string::npos) {
            start = XLCellReference(refStr.substr(0, colonPos));
            XLCellReference end(refStr.substr(colonPos + 1));
            colCount = end.column() - start.column() + 1;
        }

        std::string tableXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                               "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                               "id=\"" + std::to_string(tableId) + "\" "
                               "name=\"" + std::string(name) + "\" "
                               "displayName=\"" + std::string(name) + "\" "
                               "ref=\"" + std::string(range) + "\">"
                               "<autoFilter ref=\"" + std::string(range) + "\"/>"
                               "<tableColumns count=\"" + std::to_string(colCount) + "\">";
        for (uint32_t i = 0; i < colCount; ++i) {
            std::string h = m_worksheet->cell(start.row(), start.column() + i).value().getString();
            if (h.empty()) h = "Column" + std::to_string(i + 1);
            tableXml += "<tableColumn id=\"" + std::to_string(i + 1) + "\" name=\"" + h + "\"/>";
        }
        tableXml += "</tableColumns><tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>";
        
        m_worksheet->parentDoc().archive().addEntry(tablesFilename, tableXml);
        m_worksheet->parentDoc().contentTypes().addOverride("/" + tablesFilename, XLContentType::Table);
        
        XLXmlData* xmlData = &m_worksheet->parentDoc().m_data.emplace_back(&(m_worksheet->parentDoc()), tablesFilename, "", XLContentType::Table);
        m_loaded = false; // Force reload
        return XLTable(xmlData);
    }

    bool XLTableCollection::valid() const { return m_sheetNode != nullptr; }

    XLTable::XLTable(XLXmlData* xmlData) : XLXmlFile(xmlData) {}
    std::string XLTable::name() const { return xmlDocument().document_element().attribute("name").value(); }
    /**
     * @brief The TableAttributeOrder defines the mandatory physical sequence of attributes in the <table> tag.
     */
    const std::vector<std::string_view> TableAttributeOrder = {
        "id", "name", "displayName", "ref", "tableType", 
        "headerRowDxfId", "dataDxfId", "headerRowCount", 
        "insertRow", "insertRowShift", "totalsRowCount", 
        "totalsRowShown", "published", "comment"
    };

    /**
     * @brief Helper to set table attributes in the correct physical order.
     */
    void setTableAttribute(XMLNode& node, std::string_view name, std::string_view value) {
        // 1. Collect all existing attributes that are in our ordered list
        std::vector<std::pair<std::string, std::string>> attrs;
        for (auto attrName : TableAttributeOrder) {
            auto attr = node.attribute(std::string(attrName).c_str());
            if (attrName == name) {
                if (!value.empty()) attrs.push_back({std::string(attrName), std::string(value)});
            } else if (!attr.empty()) {
                attrs.push_back({std::string(attrName), attr.value()});
            }
        }

        // 2. Clear and re-add in order
        std::string xmlns = node.attribute("xmlns").value();
        node.remove_attributes();
        node.append_attribute("xmlns").set_value(xmlns.c_str());
        for (const auto& [n, v] : attrs) {
            node.append_attribute(n.c_str()).set_value(v.c_str());
        }
    }

    void XLTable::setName(std::string_view name) { XMLNode node = xmlDocument().document_element(); setTableAttribute(node, "name", name); }
    std::string XLTable::displayName() const { return xmlDocument().document_element().attribute("displayName").value(); }
    void XLTable::setDisplayName(std::string_view name) { XMLNode node = xmlDocument().document_element(); setTableAttribute(node, "displayName", name); }

    std::string XLTable::rangeReference() const { return xmlDocument().document_element().attribute("ref").value(); }
    void XLTable::setRangeReference(std::string_view ref)
    {
        XMLNode docNode = xmlDocument().document_element();
        
        bool hasTotals = docNode.attribute("totalsRowShown").as_bool(false);
        std::string filterRef = std::string(ref);
        
        if (hasTotals) {
             // If setting a new range while totals are shown, the autoFilter should be one row less
             auto colonPos = filterRef.find(':');
             if (colonPos != std::string::npos) {
                 XLCellReference start(filterRef.substr(0, colonPos));
                 XLCellReference end(filterRef.substr(colonPos + 1));
                 end.setRow(end.row() > 1 ? end.row() - 1 : 1);
                 filterRef = start.address() + ":" + end.address();
             }
        }
        
        setTableAttribute(docNode, "ref", ref);
        XMLNode autoFilter = appendAndGetNode(docNode, "autoFilter", TableNodeOrder);
        appendAndSetAttribute(autoFilter, "ref", filterRef);
    }

    std::string XLTable::styleName() const { return xmlDocument().document_element().child("tableStyleInfo").attribute("name").value(); }
    void XLTable::setStyleName(std::string_view styleName) { 
        XMLNode docNode = xmlDocument().document_element();
        XMLNode info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
        appendAndSetAttribute(info, "name", std::string(styleName)); 
    }

    std::string XLTable::comment() const { return ""; }
    void XLTable::setComment(std::string_view /*comment*/) {}
    
    bool XLTable::published() const { return false; }
    void XLTable::setPublished(bool /*published*/) {}

    bool XLTable::showRowStripes() const { return xmlDocument().document_element().child("tableStyleInfo").attribute("showRowStripes").as_bool(); }
    void XLTable::setShowRowStripes(bool show) { 
        XMLNode docNode = xmlDocument().document_element();
        XMLNode info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
        appendAndSetAttribute(info, "showRowStripes", show ? "1" : "0"); 
    }

    bool XLTable::showColumnStripes() const { return xmlDocument().document_element().child("tableStyleInfo").attribute("showColumnStripes").as_bool(); }
    void XLTable::setShowColumnStripes(bool show) { 
        XMLNode docNode = xmlDocument().document_element();
        XMLNode info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
        appendAndSetAttribute(info, "showColumnStripes", show ? "1" : "0"); 
    }

    bool XLTable::showFirstColumn() const { return xmlDocument().document_element().child("tableStyleInfo").attribute("showFirstColumn").as_bool(); }
    void XLTable::setShowFirstColumn(bool show) { 
        XMLNode docNode = xmlDocument().document_element();
        XMLNode info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
        appendAndSetAttribute(info, "showFirstColumn", show ? "1" : "0"); 
    }

    bool XLTable::showLastColumn() const { return xmlDocument().document_element().child("tableStyleInfo").attribute("showLastColumn").as_bool(); }
    void XLTable::setShowLastColumn(bool show) { 
        XMLNode docNode = xmlDocument().document_element();
        XMLNode info = appendAndGetNode(docNode, "tableStyleInfo", TableNodeOrder);
        appendAndSetAttribute(info, "showLastColumn", show ? "1" : "0"); 
    }

    XLAutoFilter XLTable::autoFilter() const {
        XMLNode docNode = xmlDocument().document_element();
        XMLNode filterNode = appendAndGetNode(docNode, "autoFilter", TableNodeOrder);
        return XLAutoFilter(filterNode);
    }

    XLTableColumn XLTable::appendColumn(std::string_view name) {
        XMLNode docNode = xmlDocument().document_element();
        XMLNode columns = docNode.child("tableColumns");
        uint32_t count = columns.attribute("count").as_uint(0) + 1;
        appendAndSetAttribute(columns, "count", std::to_string(count));
        
        auto column = columns.append_child("tableColumn");
        column.append_attribute("id").set_value(count);
        column.append_attribute("name").set_value(std::string(name).c_str());
        return XLTableColumn(column);
    }

    bool XLTable::showHeaderRow() const { 
        auto attr = xmlDocument().document_element().attribute("headerRowCount");
        return attr.empty() || attr.as_uint() > 0;
    }
    
    void XLTable::setShowHeaderRow(bool show) { 
        XMLNode node = xmlDocument().document_element();
        setTableAttribute(node, "headerRowCount", show ? "1" : "0");
    }

    bool XLTable::showTotalsRow() const { return xmlDocument().document_element().attribute("totalsRowShown").as_bool(); }
    void XLTable::setShowTotalsRow(bool show) {
        XMLNode node = xmlDocument().document_element();
        bool currentlyShowing = node.attribute("totalsRowShown").as_bool(false);
        
        if (show != currentlyShowing) {
            setTableAttribute(node, "totalsRowShown", show ? "1" : "0");
            setTableAttribute(node, "totalsRowCount", show ? "1" : "0");
            
            // Adjust the overall table ref, but NOT the autoFilter ref.
            // When show is true, we add 1 to the end row.
            // When show is false, we subtract 1 from the end row.
            std::string ref = node.attribute("ref").value();
            auto colonPos = ref.find(':');
            if (colonPos != std::string::npos) {
                XLCellReference start(ref.substr(0, colonPos));
                XLCellReference end(ref.substr(colonPos + 1));
                
                if (show) {
                    end.setRow(end.row() + 1);
                } else {
                    end.setRow(end.row() > 1 ? end.row() - 1 : 1);
                }
                setTableAttribute(node, "ref", start.address() + ":" + end.address());
            }
        }
    }

    void XLTable::createColumnsFromRange(const XLWorksheet& worksheet) {
        auto docNode = xmlDocument().document_element();
        auto columns = docNode.child("tableColumns");
        if (!columns.empty()) docNode.remove_child(columns);
        columns = appendAndGetNode(docNode, "tableColumns", TableNodeOrder);
        
        XLCellReference start(rangeReference().substr(0, rangeReference().find(':')));
        XLCellReference end(rangeReference().substr(rangeReference().find(':') + 1));
        uint32_t colCount = end.column() - start.column() + 1;
        appendAndSetAttribute(columns, "count", std::to_string(colCount));

        for (uint16_t c = start.column(); c <= end.column(); ++c) {
            auto column = columns.append_child("tableColumn");
            column.append_attribute("id").set_value(c - start.column() + 1);
            std::string h = worksheet.cell(start.row(), c).value().getString();
            if (h.empty()) h = "Column" + std::to_string(c - start.column() + 1);
            column.append_attribute("name").set_value(h.c_str());
        }
    }

    void XLTable::resizeToFitData(const XLWorksheet& worksheet) {
        XLCellReference start(rangeReference().substr(0, rangeReference().find(':')));
        XLCellReference end = worksheet.lastCell();
        if (end.row() < start.row()) end.setRow(start.row());
        if (end.column() < start.column()) end.setColumn(start.column());
        setRangeReference(start.address() + ":" + end.address());
    }

    XLTableColumn XLTable::column(std::string_view name) const {
        auto col = xmlDocument().document_element().child("tableColumns").find_child_by_attribute("tableColumn", "name", std::string(name).c_str());
        return XLTableColumn(col);
    }

    XLTableColumn XLTable::column(uint32_t id) const {
        auto col = xmlDocument().document_element().child("tableColumns").find_child_by_attribute("tableColumn", "id", std::to_string(id).c_str());
        return XLTableColumn(col);
    }

    void XLTable::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

} // namespace OpenXLSX

namespace OpenXLSX {
    XLTable XLTableCollection::add(std::string_view name, const XLCellRange& range) {
        return add(name, range.address());
    }
}
