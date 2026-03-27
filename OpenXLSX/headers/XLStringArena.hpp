#ifndef OPENXLSX_XLSTRINGARENA_HPP
#define OPENXLSX_XLSTRINGARENA_HPP

#include <algorithm>
#include <cstdint>
#include <memory>
#include <string_view>
#include <vector>

namespace OpenXLSX {

/**
 * @brief String memory pool (Arena Allocator) with block recycling.
 * @details Adheres to C++ Core Guidelines:
 *          1. RAII resource management (no naked new/delete)
 *          2. No copy construction/assignment to prevent dangling pointers (Rule of 5)
 *          3. Provides zero-copy access via std::string_view
 *          4. clear() recycles memory blocks instead of freeing them, eliminating
 *             malloc/free churn when the arena is reused across save cycles.
 */
class XLStringArena {
public:
    // Default block size: 4 MB accommodates many short strings without fragmentation.
    explicit XLStringArena(size_t blockSize = 4 * 1024 * 1024)
        : m_blockSize(blockSize), m_currentOffset(blockSize)
    {}

    // Disable copy — string_view pointers would dangle after a copy
    XLStringArena(const XLStringArena&)            = delete;
    XLStringArena& operator=(const XLStringArena&) = delete;

    XLStringArena(XLStringArena&&) noexcept            = default;
    XLStringArena& operator=(XLStringArena&&) noexcept = default;

    /**
     * @brief Store a string in the arena and return a non-owning view.
     * @param str The string view to store (null-terminated in the arena).
     * @return std::string_view pointing into the arena's storage.
     */
    [[nodiscard]] std::string_view store(std::string_view str)
    {
        if (str.empty()) {
            static constexpr char emptyStr[] = "";
            return {emptyStr, 0};
        }

        const size_t needed = str.size() + 1;    // +1 for null terminator

        if (m_currentOffset + needed > blockCapacity()) {
            // Try to reuse a free block, otherwise allocate a new one
            const size_t allocSize = std::max(m_blockSize, needed);
            if (!m_freeBlocks.empty() && m_freeBlocks.back().capacity >= allocSize) {
                m_activeBlocks.push_back(std::move(m_freeBlocks.back()));
                m_freeBlocks.pop_back();
            }
            else {
                m_activeBlocks.push_back({std::make_unique<char[]>(allocSize), allocSize});
            }
            m_currentOffset = 0;
        }

        char* dest = m_activeBlocks.back().data.get() + m_currentOffset;
        std::copy(str.begin(), str.end(), dest);
        dest[str.size()] = '\0';

        m_currentOffset += needed;
        return {dest, str.size()};
    }

    /**
     * @brief Recycle all blocks for reuse without freeing heap memory.
     * @details Moves active blocks to the free list so the next store() calls
     *          can reuse them without malloc. Call reset() at shutdown to
     *          actually release all memory.
     */
    void clear() noexcept
    {
        for (auto& blk : m_activeBlocks)
            m_freeBlocks.push_back(std::move(blk));
        m_activeBlocks.clear();
        m_currentOffset = m_blockSize;    // force new-block allocation on next store
    }

    /**
     * @brief Release all memory (active + free blocks). Use at shutdown.
     */
    void reset() noexcept
    {
        m_activeBlocks.clear();
        m_freeBlocks.clear();
        m_currentOffset = m_blockSize;
    }

    /**
     * @brief Total bytes allocated (active blocks only).
     */
    [[nodiscard]] size_t capacity() const noexcept
    {
        size_t total = 0;
        for (const auto& blk : m_activeBlocks) total += blk.capacity;
        return total;
    }

    /**
     * @brief Bytes used in the current active block.
     */
    [[nodiscard]] size_t used() const noexcept { return m_currentOffset; }

private:
    struct Block {
        std::unique_ptr<char[]> data;
        size_t                  capacity{0};
    };

    [[nodiscard]] size_t blockCapacity() const noexcept
    {
        return m_activeBlocks.empty() ? 0 : m_activeBlocks.back().capacity;
    }

    size_t        m_blockSize;       // target size for new blocks
    size_t        m_currentOffset;   // write position within the current active block
    std::vector<Block> m_activeBlocks;   // currently in-use blocks
    std::vector<Block> m_freeBlocks;     // recycled blocks awaiting reuse
};

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLSTRINGARENA_HPP