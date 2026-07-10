#include "OpenXLSX.hpp"
#include "TestHelpers.hpp"
#include "XLStreamWriter.hpp"
#include "XLConstants.hpp"

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>
#include <cmath>
#include <filesystem>
#include <limits>
#include <sstream>
#include <string>
#include <vector>

using namespace OpenXLSX;

namespace {
std::string uniqueXlsx(const char* tag)
{
    return OpenXLSX::TestHelpers::getUniqueFilename(tag) + ".xlsx";
}
}    // namespace

// ─────────────────────────────────────────────────────────────────────────────
// W1 — leading / trailing spaces preserved (xml:space="preserve")
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_LeadingTrailingSpaces", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_spaces");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        stream.appendRow({XLCellValue(" characters"), XLCellValue("tail "), XLCellValue(" mid ")});
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().get<std::string>() == " characters");
        REQUIRE(wks.cell(1, 2).value().get<std::string>() == "tail ");
        REQUIRE(wks.cell(1, 3).value().get<std::string>() == " mid ");

        std::stringstream ss;
        wks.cell(1, 1).print(ss);
        REQUIRE(ss.str().find("xml:space=\"preserve\"") != std::string::npos);

        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// W2 — row height / outlineLevel validation
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_RowOptsValidation", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_rowopts_val");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();

    XLStreamRowOpts tooTall;
    tooTall.height = 410.0;
    REQUIRE_THROWS_AS(stream.appendRow({XLCellValue(1)}, tooTall), XLInputError);

    XLStreamRowOpts neg;
    neg.height = -1.0;
    REQUIRE_THROWS_AS(stream.appendRow({XLCellValue(1)}, neg), XLInputError);

    XLStreamRowOpts badOutline;
    badOutline.outlineLevel = 8;
    REQUIRE_THROWS_AS(stream.appendRow({XLCellValue(1)}, badOutline), XLInputError);

    // valid boundary values should succeed
    XLStreamRowOpts ok;
    ok.height       = 409.0;
    ok.outlineLevel = 7;
    REQUIRE_NOTHROW(stream.appendRow({XLCellValue(1)}, ok));

    stream.close();
    doc.close();
    std::filesystem::remove(filename);
}

// ─────────────────────────────────────────────────────────────────────────────
// W3 — outlineLevel + hidden round-trip
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_OutlineAndHiddenRoundTrip", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_outline_hidden");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();

        XLStreamRowOpts r1;
        r1.outlineLevel = 1;
        r1.height       = 18.0;
        stream.appendRow({XLCellValue("L1")}, r1);

        XLStreamRowOpts r2;
        r2.outlineLevel = 7;
        r2.hidden       = true;
        stream.appendRow({XLCellValue("L7")}, r2);

        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.row(1).outlineLevel() == 1);
        REQUIRE(wks.row(1).height() == Catch::Approx(18.0));
        REQUIRE_FALSE(wks.row(1).isHidden());
        REQUIRE(wks.row(2).outlineLevel() == 7);
        REQUIRE(wks.row(2).isHidden());
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// W4 — sparse Empty cells: leading empties omit <c>, only later columns written
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_SparseEmptyCells", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_sparse_empty");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        stream.appendRow({XLCellValue(), XLCellValue(), XLCellValue("foo")});
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().type() == XLValueType::Empty);
        REQUIRE(wks.cell(1, 2).value().type() == XLValueType::Empty);
        REQUIRE(wks.cell(1, 3).value().get<std::string>() == "foo");

        // XML should not contain empty c for A1/B1 if skipped; C1 must exist
        std::stringstream ss;
        wks.cell(1, 3).print(ss);
        REQUIRE(ss.str().find("r=\"C1\"") != std::string::npos);

        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// W5 — RichText multi-run stream write
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_RichTextRoundTrip", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_richtext");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");

        XLRichText rt;
        XLRichTextRun run1("Rich ");
        run1.setBold(true).setFontColor(XLColor("2354E8"));
        XLRichTextRun run2("Text");
        run2.setItalic(true).setFontColor(XLColor("E83723"));
        rt.addRun(run1);
        rt.addRun(run2);

        auto stream = wks.streamWriter();
        stream.appendRow({XLCellValue(rt), XLCellValue("plain")});
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        // May round-trip as RichText or String depending on DOM load path
        const auto t = wks.cell(1, 1).value().type();
        if (t == XLValueType::RichText) {
            REQUIRE(wks.cell(1, 1).value().get<XLRichText>().plainText() == "Rich Text");
            REQUIRE(wks.cell(1, 1).value().get<XLRichText>().runs().size() >= 2);
        }
        else {
            const std::string s = wks.cell(1, 1).value().getString();
            REQUIRE(s.find("Rich") != std::string::npos);
            REQUIRE(s.find("Text") != std::string::npos);
        }
        REQUIRE(wks.cell(1, 2).value().get<std::string>() == "plain");
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// W6 — row/col range validation
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_RowColBounds", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_bounds");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();

    REQUIRE_THROWS_AS(stream.setRow(0, 1, std::vector<XLCellValue>{XLCellValue(1)}), XLInputError);
    REQUIRE_THROWS_AS(stream.setRow(MAX_ROWS + 1, 1, std::vector<XLCellValue>{XLCellValue(1)}), XLInputError);
    REQUIRE_THROWS_AS(stream.setRow(1, 0, std::vector<XLCellValue>{XLCellValue(1)}), XLInputError);
    REQUIRE_THROWS_AS(stream.setRow(1, static_cast<uint16_t>(MAX_COLS + 1), std::vector<XLCellValue>{XLCellValue(1)}), XLInputError);

    // Valid max-ish coordinates (not writing to XFD max col with huge data)
    REQUIRE_NOTHROW(stream.setRow(1, 1, std::vector<XLCellValue>{XLCellValue("ok")}));

    stream.close();
    doc.close();
    std::filesystem::remove(filename);
}

// ─────────────────────────────────────────────────────────────────────────────
// W7 — invalid cell reference
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_InvalidCellRef", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_badref");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();

    REQUIRE_THROWS(stream.setRow("A", std::vector<XLCellValue>{XLCellValue(1)}));
    REQUIRE_THROWS(stream.setRow("1A", std::vector<XLCellValue>{XLCellValue(1)}));
    // Empty address defaults to A1 in XLCellReference — not an error; first write uses row 1.
    REQUIRE_NOTHROW(stream.setRow("", std::vector<XLCellValue>{XLCellValue("via_empty_ref")}));
    REQUIRE(stream.lastRow() == 1);

    stream.close();
    doc.close();
    std::filesystem::remove(filename);
}

// ─────────────────────────────────────────────────────────────────────────────
// W8 — long string (32767+): must not crash; reopen readable
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_LongString", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_longstr");
    const std::string longStr(32769, 'c');

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        REQUIRE_NOTHROW(stream.appendRow({XLCellValue(longStr)}));
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto s = doc.workbook().worksheet("Sheet1").cell(1, 1).value().get<std::string>();
        REQUIRE(s.size() == longStr.size());
        REQUIRE(s == longStr);
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// W9 — NaN / Inf: XLCellValue only stores finite doubles → Empty
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_NonFiniteDoubles", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_nan");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();

        // Construction via XLCellValue(double): non-finite → Error("#NUM!") (library contract)
        XLCellValue nanVal(std::numeric_limits<double>::quiet_NaN());
        XLCellValue pinf(std::numeric_limits<double>::infinity());
        XLCellValue ninf(-std::numeric_limits<double>::infinity());
        REQUIRE(nanVal.type() == XLValueType::Error);
        REQUIRE(pinf.type() == XLValueType::Error);
        REQUIRE(ninf.type() == XLValueType::Error);

        stream.appendRow({nanVal, pinf, ninf, XLCellValue(1.5)});
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().type() == XLValueType::Error);
        REQUIRE(wks.cell(1, 2).value().type() == XLValueType::Error);
        REQUIRE(wks.cell(1, 3).value().type() == XLValueType::Error);
        REQUIRE(wks.cell(1, 4).value().get<double>() == Catch::Approx(1.5));
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// W10 — appendRow / setRow session row-number state machine
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_AppendThenSetRowOrder", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_row_order");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();

        stream.appendRow({XLCellValue("r1")});    // row 1
        stream.appendRow({XLCellValue("r2")});    // row 2
        REQUIRE(stream.lastRow() == 2);

        // cannot go back
        REQUIRE_THROWS_AS(stream.setRow(2, 1, std::vector<XLCellValue>{XLCellValue("x")}), XLInputError);
        REQUIRE_THROWS_AS(stream.setRow(1, 1, std::vector<XLCellValue>{XLCellValue("x")}), XLInputError);

        // can jump forward
        stream.setRow(5, 1, {XLCellValue("r5")});
        REQUIRE(stream.lastRow() == 5);

        // append continues at last+1
        stream.appendRow({XLCellValue("r6")});
        REQUIRE(stream.lastRow() == 6);

        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().get<std::string>() == "r1");
        REQUIRE(wks.cell(2, 1).value().get<std::string>() == "r2");
        REQUIRE(wks.cell(5, 1).value().get<std::string>() == "r5");
        REQUIRE(wks.cell(6, 1).value().get<std::string>() == "r6");
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Extra: formula with XML-special chars + formula-only cell
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_FormulaEscapeAndFormulaOnly", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_formula_esc");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        // formula body with characters that need escaping if ever embedded poorly
        stream.appendRow({
            XLStreamCell::withFormula("IF(A1>0,\"a&b\",\"c<d\")", XLCellValue("a&b")),
            XLStreamCell::withFormula("1+1"),    // formula only, no cache
        });
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).hasFormula());
        REQUIRE(wks.cell(1, 1).formula().get().find("a&b") != std::string::npos);
        REQUIRE(wks.cell(1, 2).hasFormula());
        REQUIRE(wks.cell(1, 2).formula().get().find("1+1") != std::string::npos);
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Extra: special XML chars in plain string
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_XmlSpecialChars", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_xmlchars");
    const std::string payload  = "a<b>&\"'c";

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        stream.appendRow({XLCellValue(payload)});
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        REQUIRE(doc.workbook().worksheet("Sheet1").cell(1, 1).value().get<std::string>() == payload);
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Extra: DOM header + stream body dimension covers end
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_HybridDomStreamDimension", "[XLStreamWriter][Edge]")
{
    const std::string filename = uniqueXlsx("stream_edge_hybrid_dim");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks               = doc.workbook().worksheet("Sheet1");
        wks.cell("A1").value() = "Header";
        wks.cell("B1").value() = "H2";
        {
            auto stream = wks.streamWriter();
            for (int i = 0; i < 50; ++i) stream.appendRow({XLCellValue(i), XLCellValue(i * 2)});
            stream.close();
        }
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().get<std::string>() == "Header");
        REQUIRE(wks.cell(51, 1).value().get<int>() == 49);
        REQUIRE(wks.rowCount() == 51);
        auto dim = wks.xmlDocument().document_element().child("dimension");
        REQUIRE_FALSE(dim.empty());
        std::string ref = dim.attribute("ref").value();
        REQUIRE(ref.find("51") != std::string::npos);
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// P2 W12 — same-session stream write is not visible to stream reader (ZIP source)
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_SameSessionReaderSeesArchiveNotTemp", "[XLStreamWriter][Edge][P2]")
{
    const std::string filename = uniqueXlsx("stream_edge_same_session");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks = doc.workbook().worksheet("Sheet1");
    // Seed archive with one row via DOM so the package has a sheet entry
    wks.cell("A1").value() = "seed";
    doc.save();

    // Stream-write more rows into temp (not yet in ZIP until save after close)
    {
        auto stream = wks.streamWriter();
        stream.appendRow({XLCellValue("streamed")});    // continues after DOM rowCount
        stream.close();
    }

    // Reader pulls from ZIP archive, not the stream temp file → still old content only
    {
        auto reader = wks.streamReader();
        REQUIRE(reader.hasNext());
        auto row = reader.nextRow();
        REQUIRE(row[0].get<std::string>() == "seed");
        // Must not see "streamed" until save+reopen
        while (reader.hasNext()) {
            auto r = reader.nextRow();
            for (const auto& c : r) {
                if (c.type() == XLValueType::String) { REQUIRE(c.get<std::string>() != "streamed"); }
            }
        }
        reader.close();
    }

    doc.save();
    doc.close();

    // After save+reopen, streamed data is present
    {
        XLDocument doc2;
        doc2.open(filename);
        auto wks2 = doc2.workbook().worksheet("Sheet1");
        // Header/seed may be row 1; streamed append starts at rowCount+1 at stream open time
        bool found = false;
        auto reader = wks2.streamReader();
        while (reader.hasNext()) {
            auto r = reader.nextRow();
            for (const auto& c : r) {
                if (c.type() == XLValueType::String && c.get<std::string>() == "streamed") found = true;
            }
        }
        reader.close();
        REQUIRE(found);
        doc2.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// P2 — error cell via stream writer round-trip
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_ErrorValueWrite", "[XLStreamWriter][Edge][P2]")
{
    const std::string filename = uniqueXlsx("stream_edge_error_write");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks = doc.workbook().worksheet("Sheet1");
        XLCellValue err;
        err.setError("#N/A");
        auto stream = wks.streamWriter();
        stream.appendRow({err, XLCellValue(1)});
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().type() == XLValueType::Error);
        REQUIRE(wks.cell(1, 2).value().get<int>() == 1);
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// P2 — append/set after close throws
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_MutateAfterCloseThrows", "[XLStreamWriter][Edge][P2]")
{
    const std::string filename = uniqueXlsx("stream_edge_after_close");

    XLDocument doc;
    doc.create(filename, XLForceOverwrite);
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto stream = wks.streamWriter();
    stream.appendRow({XLCellValue(1)});
    stream.close();
    REQUIRE_FALSE(stream.isStreamActive());
    REQUIRE_THROWS_AS(stream.appendRow({XLCellValue(2)}), XLInternalError);
    REQUIRE_THROWS_AS(stream.setRow(3, 1, std::vector<XLCellValue>{XLCellValue(3)}), XLInternalError);
    doc.close();
    std::filesystem::remove(filename);
}

// ─────────────────────────────────────────────────────────────────────────────
// P2 — setRow starting mid-column (sparse start)
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_SetRowStartColumnSparse", "[XLStreamWriter][Edge][P2]")
{
    const std::string filename = uniqueXlsx("stream_edge_start_col");

    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        stream.setRow("C1", {XLCellValue(10), XLCellValue(20)});    // A,B empty
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks = doc.workbook().worksheet("Sheet1");
        REQUIRE(wks.cell(1, 1).value().type() == XLValueType::Empty);
        REQUIRE(wks.cell(1, 2).value().type() == XLValueType::Empty);
        REQUIRE(wks.cell(1, 3).value().get<int>() == 10);
        REQUIRE(wks.cell(1, 4).value().get<int>() == 20);
        doc.close();
        std::filesystem::remove(filename);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// P2 — row style index on stream row opts
// ─────────────────────────────────────────────────────────────────────────────
TEST_CASE("StreamWriterEdge_RowStyleIndex", "[XLStreamWriter][Edge][P2]")
{
    const std::string filename = uniqueXlsx("stream_edge_row_style");

    XLStyleIndex styleId = 0;
    {
        XLDocument doc;
        doc.create(filename, XLForceOverwrite);
        XLStyle s;
        s.font.bold = true;
        styleId     = doc.styles().findOrCreateStyle(s);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto stream = wks.streamWriter();
        XLStreamRowOpts opts;
        opts.styleIndex = styleId;
        stream.appendRow({XLCellValue("x")}, opts);
        stream.close();
        doc.save();
        doc.close();
    }

    {
        XLDocument doc;
        doc.open(filename);
        auto wks    = doc.workbook().worksheet("Sheet1");
        auto reader = wks.streamReader();
        (void)reader.nextRow();
        auto ro = reader.currentRowOpts();
        REQUIRE(ro.styleIndex.has_value());
        REQUIRE(ro.styleIndex.value() == styleId);
        reader.close();
        doc.close();
        std::filesystem::remove(filename);
    }
}
