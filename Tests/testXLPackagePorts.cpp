#include <OpenXLSX.hpp>
#include <XLPackagePartFactory.hpp>
#include <XLPackageServices.hpp>
#include <XLSharedStringTable.hpp>
#include <catch2/catch_all.hpp>
#include "TestHelpers.hpp"

#include <filesystem>
#include <string>

using namespace OpenXLSX;

namespace
{
    struct TempDoc
    {
        std::string path;
        XLDocument  doc;

        explicit TempDoc(const std::string& stem)
        {
            path = TestHelpers::getUniqueFilename(stem) + ".xlsx";
            doc.create(path, XLForceOverwrite);
        }

        ~TempDoc()
        {
            if (doc.isOpen()) doc.close();
            std::error_code ec;
            std::filesystem::remove(path, ec);
        }
    };
}    // namespace

TEST_CASE("XLXmlFile::sharedStrings convenience matches document SST", "[XLPackagePorts][SharedStrings]")
{
    TempDoc tmp("sst_convenience");
    auto    wks = tmp.doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "hello-port";
    wks.cell("B1").value() = "hello-port";    // shared

    const XLSharedStrings& viaWks = wks.sharedStrings();
    const XLSharedStrings& viaDoc = tmp.doc.sharedStrings();
    REQUIRE(&viaWks == &viaDoc);

    REQUIRE(viaWks.stringCount() >= 1);
    REQUIRE(viaWks.stringExists("hello-port"));
}

TEST_CASE("XLSharedStringTable read-only port via worksheet", "[XLPackagePorts][SharedStringTable]")
{
    TempDoc tmp("sst_table");
    auto    wks = tmp.doc.workbook().worksheet("Sheet1");
    wks.cell("A1").value() = "alpha";
    wks.cell("A2").value() = "beta";

    const XLSharedStringTable& table = wks.sharedStringTable();
    REQUIRE(table.stringCount() >= 2);
    REQUIRE(table.stringExists("alpha"));
    REQUIRE(table.stringExists("beta"));

    const int32_t idx = table.getStringIndex("alpha");
    REQUIRE(idx >= 0);
    REQUIRE(std::string(table.getString(idx)) == "alpha");
    REQUIRE(table.getStringView(idx) == "alpha");

    // Document is-a SharedStringTable
    const XLSharedStringTable& asDoc = tmp.doc.sharedStringTable();
    REQUIRE(&asDoc == &static_cast<const XLSharedStringTable&>(tmp.doc));
    REQUIRE(asDoc.stringExists("beta"));
}

TEST_CASE("XLPackageServices port: archive and contentTypes", "[XLPackagePorts][PackageServices]")
{
    TempDoc tmp("pkg_services");
    auto wks = tmp.doc.workbook().worksheet("Sheet1");

    XLPackageServices& pkg = wks.package();
    REQUIRE(pkg.archive().isValid());

    // Content types knows about the workbook part
    auto& ct = pkg.contentTypes();
    bool  sawWorkbook = false;
    for (const auto& item : ct.getContentItems()) {
        if (item.type() == XLContentType::Workbook) {
            sawWorkbook = true;
            break;
        }
    }
    REQUIRE(sawWorkbook);

    // Managed part lookup
    REQUIRE(pkg.findXmlPart("xl/workbook.xml", /*doNotThrow=*/true) != nullptr);
    REQUIRE(pkg.findXmlPart("xl/does-not-exist.xml", /*doNotThrow=*/true) == nullptr);
}

TEST_CASE("XLPackagePartFactory port: nextTableId and createChart", "[XLPackagePorts][PartFactory]")
{
    TempDoc tmp("pkg_factory");
    auto    wks = tmp.doc.workbook().worksheet("Sheet1");

    XLPackagePartFactory& factory = wks.partFactory();
    // nextTableId is max(existing)+1 (does not allocate); consecutive calls agree until a table is created.
    REQUIRE(factory.nextTableId() == 1);
    REQUIRE(factory.nextTableId() == factory.nextTableId());

    XLChart chart = factory.createChart(XLChartType::Bar);
    REQUIRE(chart.getXmlPath().find("xl/charts/chart") != std::string::npos);
    REQUIRE_FALSE(chart.getXmlPath().empty());
}
