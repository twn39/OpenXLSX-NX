#include "XLImageParser.hpp"
#include "XLException.hpp"
#include <fstream>
#include <vector>

namespace OpenXLSX
{

    XLImageSize XLImageParser::parseDimensions(const std::string& path)
    {
        XLImageSize   size;
        std::ifstream file(path, std::ios::binary);
        if (!file.is_open()) { return size; }

        // Read magic numbers
        std::vector<unsigned char> magic(8);
        if (!file.read(reinterpret_cast<char*>(magic.data()), 8)) { return size; }

        // Check PNG
        if (magic[0] == 0x89 && magic[1] == 0x50 && magic[2] == 0x4E && magic[3] == 0x47) {
            size.extension = "png";
            file.seekg(16, std::ios::beg);    // Skip 8 byte signature + 8 byte IHDR chunk header
            uint32_t w = 0, h = 0;
            if (file.read(reinterpret_cast<char*>(&w), 4) && file.read(reinterpret_cast<char*>(&h), 4)) {
                // Big endian to little endian
                size.width  = (w >> 24) | ((w << 8) & 0x00FF0000) | ((w >> 8) & 0x0000FF00) | (w << 24);
                size.height = (h >> 24) | ((h << 8) & 0x00FF0000) | ((h >> 8) & 0x0000FF00) | (h << 24);
                size.valid  = true;
            }
            return size;
        }

        // Check JPG
        if (magic[0] == 0xFF && magic[1] == 0xD8) {
            size.extension = "jpeg";
            file.seekg(2, std::ios::beg);
            while (true) {
                unsigned char marker[2];
                if (!file.read(reinterpret_cast<char*>(marker), 2)) break;

                if (marker[0] != 0xFF) break;    // Not a valid marker

                if (marker[1] >= 0xC0 && marker[1] <= 0xC3) {    // SOF markers
                    file.seekg(3, std::ios::cur);                // Skip length and precision
                    unsigned char hbuf[2], wbuf[2];
                    if (file.read(reinterpret_cast<char*>(hbuf), 2) && file.read(reinterpret_cast<char*>(wbuf), 2)) {
                        size.height = (hbuf[0] << 8) | hbuf[1];
                        size.width  = (wbuf[0] << 8) | wbuf[1];
                        size.valid  = true;
                        break;
                    }
                }
                else {
                    uint16_t len = 0;
                    if (!file.read(reinterpret_cast<char*>(&len), 2)) break;
                    len = (len >> 8) | (len << 8);
                    file.seekg(len - 2, std::ios::cur);
                }
            }
            return size;
        }

        // Check GIF
        if (magic[0] == 'G' && magic[1] == 'I' && magic[2] == 'F') {
            size.extension = "gif";
            uint16_t w = 0, h = 0;
            // GIF is little-endian, no swap needed on x86, but let's be safe
            w = magic[6] | (magic[7] << 8);

            std::vector<unsigned char> hbuf(2);
            if (file.read(reinterpret_cast<char*>(hbuf.data()), 2)) {
                h           = hbuf[0] | (hbuf[1] << 8);
                size.width  = w;
                size.height = h;
                size.valid  = true;
            }
            return size;
        }

        return size;
    }

    XLImageSize XLImageParser::parseDimensions(gsl::span<const uint8_t> data)
    {
        XLImageSize size;
        if (data.size() < 8) return size;

        // Check PNG
        if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) {
            size.extension = "png";
            if (data.size() >= 24) {
                uint32_t w = (static_cast<uint32_t>(data[16]) << 24) |
                             (static_cast<uint32_t>(data[17]) << 16) |
                             (static_cast<uint32_t>(data[18]) << 8)  |
                             static_cast<uint32_t>(data[19]);
                uint32_t h = (static_cast<uint32_t>(data[20]) << 24) |
                             (static_cast<uint32_t>(data[21]) << 16) |
                             (static_cast<uint32_t>(data[22]) << 8)  |
                             static_cast<uint32_t>(data[23]);
                size.width  = w;
                size.height = h;
                size.valid  = true;
            }
            return size;
        }

        // Check JPG
        if (data[0] == 0xFF && data[1] == 0xD8) {
            size.extension = "jpeg";
            size_t offset = 2;
            while (offset + 4 <= data.size()) {
                if (data[offset] != 0xFF) break;    // Not a valid marker

                unsigned char marker = data[offset + 1];
                if (marker >= 0xC0 && marker <= 0xC3) {    // SOF markers
                    if (offset + 9 > data.size()) break;
                    size.height = (static_cast<uint32_t>(data[offset + 5]) << 8) | static_cast<uint32_t>(data[offset + 6]);
                    size.width  = (static_cast<uint32_t>(data[offset + 7]) << 8) | static_cast<uint32_t>(data[offset + 8]);
                    size.valid  = true;
                    break;
                }
                else {
                    uint16_t len = (static_cast<uint16_t>(data[offset + 2]) << 8) | static_cast<uint16_t>(data[offset + 3]);
                    offset += 2 + len;
                }
            }
            return size;
        }

        // Check GIF
        if (data[0] == 'G' && data[1] == 'I' && data[2] == 'F') {
            size.extension = "gif";
            if (data.size() >= 10) {
                size.width  = data[6] | (data[7] << 8);
                size.height = data[8] | (data[9] << 8);
                size.valid  = true;
            }
            return size;
        }

        return size;
    }
}    // namespace OpenXLSX
