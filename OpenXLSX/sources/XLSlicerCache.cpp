#include "XLSlicerCache.hpp"
#include "XLPivotTable.hpp"
#include <algorithm>

using namespace OpenXLSX;

std::string XLSlicerCache::name() const
{
    return xmlDocument().document_element().attribute("name").value();
}

void XLSlicerCache::setName(std::string_view name)
{
    auto root = xmlDocument().document_element();
    if (!root.attribute("name")) {
        root.append_attribute("name").set_value(std::string(name).c_str());
    } else {
        root.attribute("name").set_value(std::string(name).c_str());
    }
}

std::string XLSlicerCache::sourceName() const
{
    return xmlDocument().document_element().attribute("sourceName").value();
}

void XLSlicerCache::setSourceName(std::string_view sourceName)
{
    auto root = xmlDocument().document_element();
    if (!root.attribute("sourceName")) {
        root.append_attribute("sourceName").set_value(std::string(sourceName).c_str());
    } else {
        root.attribute("sourceName").set_value(std::string(sourceName).c_str());
    }
}

void XLSlicerCache::syncWithPivotCache(const XLPivotCacheDefinition& pivotCache, const std::vector<std::string>& selectedItems)
{
    auto root = xmlDocument().document_element();
    auto dataNode = root.child("data");
    if (dataNode.empty()) {
        dataNode = root.append_child("data");
    }
    auto tabularNode = dataNode.child("tabular");
    if (tabularNode.empty()) {
        tabularNode = dataNode.append_child("tabular");
    }

    std::string targetFieldName = sourceName();
    pugi::xml_node targetField;
    auto cacheFields = pivotCache.xmlDocument().document_element().child("cacheFields");
    for (auto field : cacheFields.children("cacheField")) {
        if (targetFieldName == field.attribute("name").value()) {
            targetField = field;
            break;
        }
    }

    if (targetField.empty()) return;

    auto sharedItemsNode = targetField.child("sharedItems");
    if (sharedItemsNode.empty()) return;

    auto itemsNode = tabularNode.child("items");
    if (!itemsNode.empty()) {
        tabularNode.remove_child(itemsNode);
    }
    itemsNode = tabularNode.append_child("items");

    uint32_t index = 0;
    uint32_t count = 0;
    for (auto item : sharedItemsNode.children()) {
        std::string val = item.attribute("v").value();
        
        bool isSelected = selectedItems.empty() || 
            (std::find(selectedItems.begin(), selectedItems.end(), val) != selectedItems.end());

        auto itemNode = itemsNode.append_child("i");
        itemNode.append_attribute("x").set_value(index++);
        if (!isSelected) {
            itemNode.append_attribute("s").set_value(false);
        }
        count++;
    }
    itemsNode.append_attribute("count").set_value(count);
}
