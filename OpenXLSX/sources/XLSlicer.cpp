#include "XLSlicer.hpp"

using namespace OpenXLSX;

std::string XLSlicer::name() const
{
    return xmlDocument().document_element().child("slicer").attribute("name").value();
}

void XLSlicer::setName(std::string_view name)
{
    auto slicer = xmlDocument().document_element().child("slicer");
    if (!slicer) return;
    if (!slicer.attribute("name")) {
        slicer.append_attribute("name").set_value(std::string(name).c_str());
    } else {
        slicer.attribute("name").set_value(std::string(name).c_str());
    }
}

std::string XLSlicer::caption() const
{
    return xmlDocument().document_element().child("slicer").attribute("caption").value();
}

void XLSlicer::setCaption(std::string_view caption)
{
    auto slicer = xmlDocument().document_element().child("slicer");
    if (!slicer) return;
    if (!slicer.attribute("caption")) {
        slicer.append_attribute("caption").set_value(std::string(caption).c_str());
    } else {
        slicer.attribute("caption").set_value(std::string(caption).c_str());
    }
}

std::string XLSlicer::slicerStyle() const
{
    return xmlDocument().document_element().child("slicer").attribute("slicerStyle").value();
}

void XLSlicer::setSlicerStyle(std::string_view styleName)
{
    auto slicer = xmlDocument().document_element().child("slicer");
    if (!slicer) return;
    if (!slicer.attribute("slicerStyle")) {
        slicer.append_attribute("slicerStyle").set_value(std::string(styleName).c_str());
    } else {
        slicer.attribute("slicerStyle").set_value(std::string(styleName).c_str());
    }
}

bool XLSlicer::showCaption() const
{
    auto slicer = xmlDocument().document_element().child("slicer");
    if (!slicer) return true;
    auto attr = slicer.attribute("showCaption");
    if (!attr) return true;
    return attr.as_bool(true);
}

void XLSlicer::setShowCaption(bool show)
{
    auto slicer = xmlDocument().document_element().child("slicer");
    if (!slicer) return;
    if (!slicer.attribute("showCaption")) {
        slicer.append_attribute("showCaption").set_value(show);
    } else {
        slicer.attribute("showCaption").set_value(show);
    }
}

std::string XLSlicer::cache() const
{
    return xmlDocument().document_element().child("slicer").attribute("cache").value();
}

void XLSlicer::setCache(std::string_view cacheName)
{
    auto slicer = xmlDocument().document_element().child("slicer");
    if (!slicer) return;
    if (!slicer.attribute("cache")) {
        slicer.append_attribute("cache").set_value(std::string(cacheName).c_str());
    } else {
        slicer.attribute("cache").set_value(std::string(cacheName).c_str());
    }
}
