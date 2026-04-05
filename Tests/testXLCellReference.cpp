#include <OpenXLSX.hpp>
#include <catch2/catch_all.hpp>
#include <fstream>

using namespace OpenXLSX;

TEST_CASE("XLCellReference Tests", "[XLCell]")
{
    SECTION("Constructors")
    {
        auto ref = XLCellReference();
        REQUIRE(ref.address() == "A1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 1);

        ref = XLCellReference("B2");
        REQUIRE(ref.address() == "B2");
        REQUIRE(ref.row() == 2);
        REQUIRE(ref.column() == 2);

        ref = XLCellReference(3, 3);
        REQUIRE(ref.address() == "C3");
        REQUIRE(ref.row() == 3);
        REQUIRE(ref.column() == 3);

        ref = XLCellReference(4, "D");
        REQUIRE(ref.address() == "D4");
        REQUIRE(ref.row() == 4);
        REQUIRE(ref.column() == 4);

        ref = XLCellReference("AA1");
        REQUIRE(ref.address() == "AA1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 27);

        ref = XLCellReference(1, 27);
        REQUIRE(ref.address() == "AA1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 27);

        ref = XLCellReference("XFD1");
        REQUIRE(ref.address() == "XFD1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == MAX_COLS);

        ref = XLCellReference("A1048576");
        REQUIRE(ref.address() == "A1048576");
        REQUIRE(ref.row() == MAX_ROWS);
        REQUIRE(ref.column() == 1);

        REQUIRE_THROWS(XLCellReference(0, 0));
        REQUIRE_THROWS(XLCellReference("XFE1"));
        REQUIRE_THROWS(XLCellReference("A1048577"));
        REQUIRE_THROWS(XLCellReference(MAX_ROWS + 1, MAX_COLS + 1));
        REQUIRE_THROWS(XLCellReference(MAX_ROWS + 1, "XFE"));

        ref        = XLCellReference("B2");
        auto c_ref = ref;
        REQUIRE(c_ref.address() == "B2");
        REQUIRE(c_ref.row() == 2);
        REQUIRE(c_ref.column() == 2);

        ref   = XLCellReference("C3");
        c_ref = ref;
        REQUIRE(c_ref.address() == "C3");
        REQUIRE(c_ref.row() == 3);
        REQUIRE(c_ref.column() == 3);

        ref        = XLCellReference("B2");
        auto m_ref = std::move(ref);
        REQUIRE(m_ref.address() == "B2");
        REQUIRE(m_ref.row() == 2);
        REQUIRE(m_ref.column() == 2);

        ref   = XLCellReference("C3");
        m_ref = std::move(ref);
        REQUIRE(m_ref.address() == "C3");
        REQUIRE(m_ref.row() == 3);
        REQUIRE(m_ref.column() == 3);
    }

    SECTION("Data-driven Reference Parsing")
    {
        auto [address, expectedRow, expectedCol, isValid] = GENERATE(table<std::string, uint32_t, uint16_t, bool>({
            // Valid common cases
            {"A1", 1, 1, true},
            {"Z10", 10, 26, true},
            {"AA100", 100, 27, true},
            {"XFD1048576", 1048576, 16384, true},
            // Absolute references (should strip $)
            {"$A$1", 1, 1, false},
            {"$B2", 2, 2, false},
            {"C$3", 3, 3, false},
            // Invalid references
            {"1", 0, 0, false},
            {"A", 0, 0, false},
            {"A0", 0, 0, false},
            {"1A", 0, 0, false},
            {"A-1", 0, 0, false},
            {"XFE1", 0, 0, false},   // Beyond XFD
            {"A1048577", 0, 0, false} // Beyond 1048576
        }));

        if (isValid) {
            REQUIRE_NOTHROW(XLCellReference(address));
            XLCellReference ref(address);
            REQUIRE(ref.row() == expectedRow);
            REQUIRE(ref.column() == expectedCol);
        } else {
            INFO("Testing invalid address: " << address);
            REQUIRE_THROWS(XLCellReference(address));
        }
    }

    SECTION("Increment and Decrement")
    {
        auto ref = XLCellReference();
        REQUIRE(ref.address() == "A1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 1);

        ++ref;
        REQUIRE(ref.address() == "B1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 2);

        ref++;
        REQUIRE(ref.address() == "C1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 3);

        --ref;
        REQUIRE(ref.address() == "B1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 2);

        ref--;
        REQUIRE(ref.address() == "A1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 1);

        --ref;
        REQUIRE(ref.address() == "XFD1048576");
        REQUIRE(ref.row() == MAX_ROWS);
        REQUIRE(ref.column() == MAX_COLS);

        ++ref;
        REQUIRE(ref.address() == "A1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == 1);

        ref = XLCellReference(1, MAX_COLS);

        ++ref;
        REQUIRE(ref.address() == "A2");
        REQUIRE(ref.row() == 2);
        REQUIRE(ref.column() == 1);

        --ref;
        REQUIRE(ref.address() == "XFD1");
        REQUIRE(ref.row() == 1);
        REQUIRE(ref.column() == MAX_COLS);
    }

    SECTION("Comparison operators")
    {
        auto ref1 = XLCellReference("B2");
        auto ref2 = XLCellReference("B2");
        auto ref3 = XLCellReference("C3");

        REQUIRE(ref1 == ref2);
        REQUIRE_FALSE(ref1 == ref3);
        REQUIRE(ref1 != ref3);
        REQUIRE_FALSE(ref1 != ref2);
        REQUIRE(ref1 < ref3);
        REQUIRE_FALSE(ref1 > ref3);
        REQUIRE(ref3 > ref1);
        REQUIRE_FALSE(ref3 < ref1);
        REQUIRE(ref1 <= ref2);
        REQUIRE(ref1 <= ref3);
        REQUIRE_FALSE(ref3 <= ref1);
        REQUIRE(ref2 >= ref1);
        REQUIRE(ref3 >= ref1);
        REQUIRE_FALSE(ref1 >= ref3);
    }
}