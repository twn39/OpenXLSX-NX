// ===== External Includes ===== //
#include <stdexcept>
#include <pugixml.hpp>
#include <fast_float/fast_float.h>

// ===== OpenXLSX Includes ===== //
#include "XLStreamReader.hpp"
#include "XLWorksheet.hpp"
#include "XLDocument.hpp"
#include "XLSharedStrings.hpp"
#include "XLCellReference.hpp"

namespace OpenXLSX {

    XLStreamReader::XLStreamReader(const XLWorksheet* worksheet) 
        : m_worksheet(worksheet)
    {
        if (!worksheet) throw XLInternalError("Worksheet is null");
        
        m_zipStream = worksheet->parentDoc().archive().openEntryStream(worksheet->getXmlPath());
        if (!m_zipStream) throw XLInternalError("Failed to open worksheet zip stream");
    }

    XLStreamReader::~XLStreamReader() {
        cleanup();
    }

    XLStreamReader::XLStreamReader(XLStreamReader&& other) noexcept
        : m_worksheet(other.m_worksheet),
          m_zipStream(other.m_zipStream),
          m_buffer(std::move(other.m_buffer)),
          m_eof(other.m_eof),
          m_currentRow(other.m_currentRow)
    {
        other.m_worksheet = nullptr;
        other.m_zipStream = nullptr;
        other.m_eof = true;
    }

    XLStreamReader& XLStreamReader::operator=(XLStreamReader&& other) noexcept {
        if (this != &other) {
            cleanup();
            m_worksheet = other.m_worksheet;
            m_zipStream = other.m_zipStream;
            m_buffer = std::move(other.m_buffer);
            m_eof = other.m_eof;
            m_currentRow = other.m_currentRow;

            other.m_worksheet = nullptr;
            other.m_zipStream = nullptr;
            other.m_eof = true;
        }
        return *this;
    }

    void XLStreamReader::cleanup() {
        if (m_zipStream && m_worksheet) {
            m_worksheet->parentDoc().archive().closeEntryStream(m_zipStream);
            m_zipStream = nullptr;
        }
    }

    void XLStreamReader::fetchMoreData() {
        if (m_eof || !m_zipStream) return;

        char buf[65536]; // 64KB chunks
        int64_t bytesRead = m_worksheet->parentDoc().archive().readEntryStream(m_zipStream, buf, sizeof(buf));
        
        if (bytesRead > 0) {
            m_buffer.append(buf, bytesRead);
        } else {
            m_eof = true;
            cleanup();
        }
    }

    std::string XLStreamReader::extractNextRowXml() {
        while (true) {
            // Find start of a row
            size_t startPos = m_buffer.find("<row");
            if (startPos == std::string::npos) {
                if (m_eof) {
                    m_buffer.clear(); // Clear out remaining junk
                    return "";
                }
                // Clear out before fetching more to avoid buffer growing too large, 
                // but only if we know there's no "<row" prefix being formed (e.g., '<' at the very end).
                // To be safe, keep the last few characters if they might form "<row".
                if (m_buffer.size() > 5) {
                    m_buffer.erase(0, m_buffer.size() - 5);
                }
                fetchMoreData();
                continue;
            }

            // Find end of the row
            // A row can end with </row> or be empty like <row r="2"/>
            size_t endTagPos = m_buffer.find("</row>", startPos);
            size_t closeTagPos = m_buffer.find("/>", startPos);
            size_t openTagClosePos = m_buffer.find(">", startPos);

            if (openTagClosePos == std::string::npos) {
                if (m_eof) return "";
                fetchMoreData();
                continue;
            }

            // If the row tag closes with />
            if (closeTagPos != std::string::npos && closeTagPos == openTagClosePos - 1) {
                std::string rowXml = m_buffer.substr(startPos, openTagClosePos - startPos + 1);
                m_buffer.erase(0, openTagClosePos + 1);
                return rowXml;
            }

            // It's a standard <row> ... </row>
            if (endTagPos == std::string::npos) {
                if (m_eof) return "";
                fetchMoreData();
                continue;
            }

            std::string rowXml = m_buffer.substr(startPos, endTagPos + 6 - startPos);
            m_buffer.erase(0, endTagPos + 6);
            return rowXml;
        }
    }

    bool XLStreamReader::hasNext() {
        if (!m_buffer.empty() && m_buffer.find("<row") != std::string::npos) {
            return true;
        }
        
        while (!m_eof) {
            fetchMoreData();
            if (m_buffer.find("<row") != std::string::npos) {
                return true;
            }
        }
        return false;
    }

    std::vector<XLCellValue> XLStreamReader::nextRow() {
        std::vector<XLCellValue> result;
        std::string rowXml = extractNextRowXml();
        if (rowXml.empty()) return result;

        pugi::xml_document doc;
        // Optimization: parse fragment directly
        pugi::xml_parse_result parseRes = doc.load_buffer(rowXml.data(), rowXml.size());
        if (!parseRes) return result;

        pugi::xml_node rowNode = doc.child("row");
        if (!rowNode) return result;

        uint32_t expectedColIdx = 1;

        for (pugi::xml_node cNode = rowNode.child("c"); cNode; cNode = cNode.next_sibling("c")) {
            std::string rAttr = cNode.attribute("r").value();
            uint16_t actualColIdx = XLCellReference(rAttr).column();

            // Fill missing columns
            while (expectedColIdx < actualColIdx) {
                result.emplace_back(XLCellValue());
                expectedColIdx++;
            }

            std::string tAttr = cNode.attribute("t").value();
            pugi::xml_node vNode = cNode.child("v");

            if (tAttr == "s") {
                // Shared string
                if (vNode) {
                    int32_t sIdx = std::stoi(vNode.text().get());
                    result.emplace_back(std::string(m_worksheet->parentDoc().sharedStrings().getString(sIdx)));
                } else {
                    result.emplace_back(XLCellValue());
                }
            } else if (tAttr == "inlineStr") {
                // Inline string
                pugi::xml_node tNode = cNode.child("is").child("t");
                if (tNode) {
                    result.emplace_back(std::string(tNode.text().get()));
                } else {
                    result.emplace_back(XLCellValue());
                }
            } else if (tAttr == "b") {
                // Boolean
                if (vNode) {
                    bool val = (std::string_view(vNode.text().get()) == "1");
                    result.emplace_back(val);
                } else {
                    result.emplace_back(XLCellValue());
                }
            } else if (tAttr == "str") {
                // Formula string
                if (vNode) {
                    result.emplace_back(std::string(vNode.text().get()));
                } else {
                    result.emplace_back(XLCellValue());
                }
            } else if (tAttr == "e") {
                // Error
                if (vNode) {
                    XLCellValue errorVal;
                    errorVal.setError(vNode.text().get());
                    result.emplace_back(errorVal);
                } else {
                    result.emplace_back(XLCellValue());
                }
            } else {
                // Number (t="n" or empty)
                if (vNode) {
                    double val;
                    auto textStr = vNode.text().get();
                    auto [ptr, ec] = fast_float::from_chars(textStr, textStr + strlen(textStr), val);
                    if (ec == std::errc()) {
                        if (strchr(textStr, '.') != nullptr || strchr(textStr, 'E') != nullptr || strchr(textStr, 'e') != nullptr) {
                            result.emplace_back(val);
                        } else {
                            result.emplace_back(static_cast<int64_t>(val));
                        }
                    } else {
                        result.emplace_back(XLCellValue());
                    }
                } else {
                    result.emplace_back(XLCellValue());
                }
            }

            expectedColIdx++;
        }

        m_currentRow++;
        return result;
    }

}