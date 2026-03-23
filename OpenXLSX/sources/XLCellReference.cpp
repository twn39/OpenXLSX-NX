// ===== External Includes ===== //
#include <array>
#include <cmath>
#include <string_view>
#include <gsl/assert>
#include <gsl/util>
#include <charconv>
#include <cctype>     // std::isdigit
#include <cstdint>    // pull requests #216, #232

// ===== OpenXLSX Includes ===== //
#include "XLCellReference.hpp"
#include "XLConstants.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

constexpr uint8_t alphabetSize = 26;

constexpr uint8_t asciiOffset = 64;

namespace
{
    constexpr bool addressIsValid(uint32_t row, uint16_t column) noexcept
    { return !(row < 1 or row > OpenXLSX::MAX_ROWS or column < 1 or column > OpenXLSX::MAX_COLS); }
}    // namespace

/**
 * @details Initializes a reference using an Excel-style coordinate string (e.g., 'A1'). Serves as the primary parser for cell identities read directly from the DOM.
 */
XLCellReference::XLCellReference(std::string_view cellAddress)
{
    if (not cellAddress.empty()) {
        setAddress(cellAddress);
        if (not addressIsValid(m_row, m_column)) throw XLCellAddressError("Cell reference is invalid");
    }
    // else: use default values initialized in header (m_row = 1, m_column = 1, m_cellAddress = "A1")
}

/**
 * @details Directly binds explicit row and column indices. Offers higher performance than string parsing when iterating systematically (e.g., over a numerical range).
 */
XLCellReference::XLCellReference(uint32_t row, uint16_t column)
{
    if (!addressIsValid(row, column)) throw XLCellAddressError("Cell reference is invalid");
    setRowAndColumn(row, column);
}

/**
 * @details Binds a hybrid coordinate system using explicit numeric rows and alphabetical column headers, typically for convenient manual cell selection.
 */
XLCellReference::XLCellReference(uint32_t row, std::string_view column)
{
    if (!addressIsValid(row, columnAsNumber(column))) throw XLCellAddressError("Cell reference is invalid");
    setRowAndColumn(row, columnAsNumber(column));
}

/**
 * @details
 */
XLCellReference::XLCellReference(const XLCellReference& other) = default;

/**
 * @details
 */
XLCellReference::XLCellReference(XLCellReference&& other) noexcept = default;

/**
 * @details
 */
XLCellReference::~XLCellReference() = default;

/**
 * @details
 */
XLCellReference& XLCellReference::operator=(const XLCellReference& other) = default;

/**
 * @details
 */
XLCellReference& XLCellReference::operator=(XLCellReference&& other) noexcept = default;

/**
 * @details
 */
XLCellReference& XLCellReference::operator++()
{
    if (m_column < MAX_COLS) { setColumn(m_column + 1); }
    else if (m_column == MAX_COLS and m_row < MAX_ROWS) {
        m_column = 1;
        setRow(m_row + 1);
    }
    else if (m_column == MAX_COLS and m_row == MAX_ROWS) {
        m_column      = 1;
        m_row         = 1;
        m_cellAddress = "A1";
    }

    return *this;
}

/**
 * @details
 */
XLCellReference XLCellReference::operator++(int)
{    // NOLINT
    auto oldRef(*this);
    ++(*this);
    return oldRef;
}

/**
 * @details
 */
XLCellReference& XLCellReference::operator--()
{
    if (m_column > 1) { setColumn(m_column - 1); }
    else if (m_column == 1 and m_row > 1) {
        m_column = MAX_COLS;
        setRow(m_row - 1);
    }
    else if (m_column == 1 and m_row == 1) {
        m_column      = MAX_COLS;
        m_row         = MAX_ROWS;
        m_cellAddress = "XFD1048576";    // this address represents the very last cell that an excel spreadsheet can reference / support
    }
    return *this;
}

/**
 * @details
 */
XLCellReference XLCellReference::operator--(int)
{    // NOLINT
    auto oldRef(*this);
    --(*this);
    return oldRef;
}

/**
 * @details Provides the raw integer index representing the 1-based vertical position within the worksheet.
 */
uint32_t XLCellReference::row() const noexcept { return m_row; }

/**
 * @details Mutates the 1-based vertical index. Automatically recalculates the full Excel-style string address to keep the internal state in sync.
 */
void XLCellReference::setRow(uint32_t row)
{
    if (!addressIsValid(row, m_column)) throw XLCellAddressError("Cell reference is invalid");

    m_row         = row;
    m_cellAddress = columnAsString(m_column) + rowAsString(m_row);
}

/**
 * @details Provides the raw integer index representing the 1-based horizontal position within the worksheet.
 */
uint16_t XLCellReference::column() const noexcept { return m_column; }

/**
 * @details Mutates the 1-based horizontal index. Validates the boundary and keeps the cached string address in sync.
 */
void XLCellReference::setColumn(uint16_t column)
{
    if (!addressIsValid(m_row, column)) throw XLCellAddressError("Cell reference is invalid");

    m_column      = column;
    m_cellAddress = columnAsString(m_column) + rowAsString(m_row);
}

/**
 * @details Re-binds the object to an entirely new coordinate, bypassing dual string recalculation by batching the integer mutations together.
 */
void XLCellReference::setRowAndColumn(uint32_t row, uint16_t column)
{
    if (!addressIsValid(row, column)) throw XLCellAddressError("Cell reference is invalid");

    m_row         = row;
    m_column      = column;
    m_cellAddress = columnAsString(m_column) + rowAsString(m_row);
}

/**
 * @details Provides the cached Excel-style cell identifier (e.g. 'A1'), which is directly compliant with the OOXML <c r=\"A1\"> coordinate schema requirement.
 */
std::string XLCellReference::address() const { return m_cellAddress; }

/**
 * @details Parses an Excel-style coordinate string (e.g., 'A1') and breaks it down into raw 1-based vertical and horizontal integer components for internal programmatic routing.
 */
void XLCellReference::setAddress(std::string_view address)
{
    const auto [row, col] = coordinatesFromAddress(address);
    m_row                 = row;
    m_column              = col;
    m_cellAddress         = address;
}

/**
 * @details Highly optimized formatter converting integer indices to strings (using std::to_chars if available) to minimize serialization overhead.
 */
std::string XLCellReference::rowAsString(uint32_t row)
{
    std::array<char, 11> str{};    // NOLINT (11 chars enough for max uint32_t)
    const auto*         p = std::to_chars(str.data(), str.data() + str.size(), row).ptr;
    return std::string{str.data(), static_cast<size_t>(p - str.data())};
}

/**
 * @details Fast integer parser specifically meant to process the numeric half of an Excel cell address.
 */
uint32_t XLCellReference::rowAsNumber(std::string_view row)
{
    uint32_t value = 0;
    std::from_chars(row.data(), row.data() + row.size(), value);    // NOLINT
    return value;
}

/**
 * @details Translates a 1-based horizontal integer index into the standard Excel alphabetical column notation (A, B, ..., Z, AA, AB...). Required for generating compliant OOXML XML nodes.
 */
std::string XLCellReference::columnAsString(uint16_t column)
{
    Expects(column >= 1 && column <= MAX_COLS);
    std::string result;

    constexpr uint16_t oneLetterMax = alphabetSize;
    constexpr uint16_t twoLetterMax = alphabetSize + alphabetSize * alphabetSize;

    // ===== If there is one letter in the Column Name:
    if (column <= oneLetterMax) result += gsl::narrow_cast<char>(column + asciiOffset);

    // ===== If there are two letters in the Column Name:
    else if (column > oneLetterMax and column <= twoLetterMax) {
        result += gsl::narrow_cast<char>((column - (oneLetterMax + 1)) / alphabetSize + asciiOffset + 1);
        result += gsl::narrow_cast<char>((column - (oneLetterMax + 1)) % alphabetSize + asciiOffset + 1);
    }

    // ===== If there are three letters in the Column Name:
    else {
        result += gsl::narrow_cast<char>((column - (twoLetterMax + 1)) / (alphabetSize * alphabetSize) + asciiOffset + 1);
        result += gsl::narrow_cast<char>(((column - (twoLetterMax + 1)) / alphabetSize) % alphabetSize + asciiOffset + 1);
        result += gsl::narrow_cast<char>((column - (twoLetterMax + 1)) % alphabetSize + asciiOffset + 1);
    }

    return result;
}

/**
 * @details Performs Base26 reverse-translation from alphabetical column notation into a numeric index. Throws XLInputError for non-alphabet characters or violations of maximum bounds.
 */
uint16_t XLCellReference::columnAsNumber(std::string_view column)
{
    uint64_t letterCount = 0;
    uint32_t colNo       = 0;
    for (const auto letter : column) {
        if (letter >= 'A' and letter <= 'Z') {    // allow only uppercase letters
            ++letterCount;
            colNo = colNo * 26 + static_cast<uint32_t>(letter - 'A' + 1);
        }
        else
            break;
    }

    // ===== If the full string was decoded and colNo is within allowed range [1;MAX_COLS]
    if (letterCount == column.length() and colNo > 0 and colNo <= MAX_COLS) return gsl::narrow_cast<uint16_t>(colNo);
    throw XLInputError("XLCellReference::columnAsNumber - column \"" + std::string(column) + "\" is invalid");


}

/**
 * @details Splits a consolidated alphanumeric coordinate (e.g., 'A1') into discrete row and column components. Throws if the format is malformed or out of bounds.
 */
XLCoordinates XLCellReference::coordinatesFromAddress(std::string_view address)
{
    uint64_t letterCount = 0;
    uint32_t colNo       = 0;
    for (const auto letter : address) {
        if (letter >= 'A' and letter <= 'Z') {    // allow only uppercase letters
            ++letterCount;
            colNo = colNo * 26 + static_cast<uint32_t>(letter - 'A' + 1);
        }
        else
            break;
    }

    // ===== If address contains between 1 and 3 letters and has at least 1 more character for the row
    if (colNo > 0 and colNo <= MAX_COLS and address.length() > letterCount) {
        size_t   pos   = letterCount;
        uint64_t rowNo = 0;
        for (; pos < address.length() and std::isdigit(address[pos]); ++pos)    // check digits
            rowNo = rowNo * 10 + static_cast<uint64_t>(address[pos] - '0');
        if (pos == address.length() and rowNo <= MAX_ROWS)    // full address was < 4 letters + only digits
            return {gsl::narrow_cast<uint32_t>(rowNo), gsl::narrow_cast<uint16_t>(colNo)};
    }
    throw XLInputError("XLCellReference::coordinatesFromAddress - address \"" + std::string(address) + "\" is invalid");


}
