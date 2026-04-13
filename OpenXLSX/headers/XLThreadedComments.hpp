#ifndef OPENXLSX_XLTHREADEDCOMMENTS_HPP
#define OPENXLSX_XLTHREADEDCOMMENTS_HPP

#include <string>
#include <vector>
#include <gsl/pointers>
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    /**
     * @brief A proxy class encapsulating a single modern threaded comment.
     */
    class OPENXLSX_EXPORT XLThreadedComment
    {
    public:
        XLThreadedComment() = default;
        explicit XLThreadedComment(XMLNode node, XLWorksheet* wks = nullptr);
        ~XLThreadedComment() = default;
        XLThreadedComment(const XLThreadedComment& other) = default;
        XLThreadedComment(XLThreadedComment&& other) = default;
        XLThreadedComment& operator=(const XLThreadedComment& other) = default;
        XLThreadedComment& operator=(XLThreadedComment&& other) = default;

        bool valid() const;

        std::string ref() const;
        std::string id() const;
        std::string parentId() const;
        std::string personId() const;
        std::string text() const;
        
        bool isResolved() const;
        void setResolved(bool resolved);

        /**
         * @brief Add a reply to this thread (only works if initialized with Worksheet context).
         */
        XLThreadedComment& addReply(const std::string& text, const std::string& author = "");

        void setWorksheet(XLWorksheet* wks) { m_wks = wks; }

    private:
        XMLNode m_node;
        XLWorksheet* m_wks{nullptr};
    };

    class XLWorksheet;

    /**
     * @brief A class encapsulating modern Excel threaded comments (ThreadedComments.xml)
     */
    class OPENXLSX_EXPORT XLThreadedComments : public XLXmlFile
    {
    public:
        XLThreadedComments() : XLXmlFile(nullptr) {}
        explicit XLThreadedComments(gsl::not_null<XLXmlData*> xmlData) : XLXmlFile(xmlData) {}
        ~XLThreadedComments() = default;
        XLThreadedComments(const XLThreadedComments& other) = default;
        XLThreadedComments(XLThreadedComments&& other) = default;
        XLThreadedComments& operator=(const XLThreadedComments& other) = default;
        XLThreadedComments& operator=(XLThreadedComments&& other) = default;

        /**
         * @brief Get the top-level comment thread for a specific cell reference.
         */
        XLThreadedComment comment(const std::string& ref) const;

        /**
         * @brief Get all replies in a specific comment thread by parent ID.
         */
        std::vector<XLThreadedComment> replies(const std::string& parentId) const;

        /**
         * @brief Add a new top-level threaded comment.
         */
        XLThreadedComment addComment(std::string_view ref, std::string_view personId, std::string_view text);

        /**
         * @brief Add a reply to an existing threaded comment.
         */
        XLThreadedComment addReply(const std::string& parentId, const std::string& personId, const std::string& text);

        /**
         * @brief Removes a threaded comment and its replies.
         */
        bool deleteComment(const std::string& ref);
    };

    /**
     * @brief A proxy class encapsulating a single person (author) entity.
     */
    class OPENXLSX_EXPORT XLPerson
    {
    public:
        XLPerson() = default;
        explicit XLPerson(XMLNode node);
        ~XLPerson() = default;
        XLPerson(const XLPerson& other) = default;
        XLPerson(XLPerson&& other) = default;
        XLPerson& operator=(const XLPerson& other) = default;
        XLPerson& operator=(XLPerson&& other) = default;

        bool valid() const;

        std::string id() const;
        std::string displayName() const;

    private:
        XMLNode m_node;
    };

    /**
     * @brief A class encapsulating modern Excel persons metadata (persons.xml)
     */
    class OPENXLSX_EXPORT XLPersons : public XLXmlFile
    {
    public:
        XLPersons() : XLXmlFile(nullptr) {}
        explicit XLPersons(gsl::not_null<XLXmlData*> xmlData) : XLXmlFile(xmlData) {}
        ~XLPersons() = default;
        XLPersons(const XLPersons& other) = default;
        XLPersons(XLPersons&& other) = default;
        XLPersons& operator=(const XLPersons& other) = default;
        XLPersons& operator=(XLPersons&& other) = default;

        /**
         * @brief Retrieve a person entity by their unique personId.
         */
        XLPerson person(const std::string& id) const;

        /**
         * @brief Add a person by their display name, generating a new ID if it doesn't exist.
         * @return The unique personId.
         */
        std::string addPerson(const std::string& displayName);
    };
} // namespace OpenXLSX

#endif // OPENXLSX_XLTHREADEDCOMMENTS_HPP