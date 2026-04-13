# Quick Start: Cell Comments & Threads {#comments_tutorial}

OpenXLSX-NX provides robust, full-featured support for cell annotations in Microsoft Excel, supporting both the classic 97-2016 comment boxes and the modern Office 365 threaded conversations.

1. **Traditional Comments (Notes):** The classic floating yellow boxes with a red triangle indicator in the cell corner. These are typically used for static annotations or instructions.
2. **Modern Threaded Comments:** The modern conversational UI (introduced in Office 365) with a purple indicator. These allow multiple users to have a back-and-forth discussion directly tied to a cell.

OpenXLSX-NX offers a highly ergonomic **Fluent Chaining API** and **Document-Level Default Authors** to make generating these annotations seamless and zero-boilerplate.

---

## 1. Fluent Setup & Default Authors

When generating reports programmatically, you often act on behalf of a system or a single user. You can set a global default author to avoid passing the same name repeatedly.

```cpp
#include <OpenXLSX.hpp>

using namespace OpenXLSX;

int main() {
    XLDocument doc;
    doc.create("CommentsDemo.xlsx", XLForceOverwrite);
    
    // 1. Set a global default author for all generated comments/notes
    doc.setDefaultAuthor("Financial Auditor");
    
    auto wks = doc.workbook().worksheet("Sheet1");
    wks.setName("Modern Comments");
```

---

## 2. Modern Threaded Comments (Excel 365)

For modern conversational comments, OpenXLSX-NX provides highly streamlined, object-oriented methods directly on the `XLCell` object. When you create a modern comment, OpenXLSX-NX *automatically* generates the invisible legacy fallback structures required by the OOXML standard for backward compatibility with older Excel versions.

```cpp
    // 1. Start a new conversation thread on a cell fluently
    // Signature: addComment(Text, [Optional AuthorName])
    // Returns an XLThreadedComment object which you can chain replies to.
    auto thread = wks.cell("A3")
                     .value() = "Q3 Projections";
                     
    wks.cell("A3").addComment("Can we update the Q3 projections?", "Alice Smith")
                  .addReply("Yes, I will have them ready by tomorrow.", "Bob Johnson")
                  .addReply("Great, thanks Bob!", "Alice Smith")
                  .setResolved(true); // Mark the entire thread as resolved (Excel 365 feature)
```

---

## 3. Traditional Comments (Notes)

If you strictly need the classic yellow boxes, you can use the `addNote()` fluent method. You also have full access to the underlying `XLComments` system to manipulate the VML drawing shapes (e.g., resizing the box or making it permanently visible).

```cpp
    // 1. Add a classic yellow note to a cell
    wks.cell("C2").value() = "Audit Target";
    wks.cell("C2").addNote("Please review these numbers."); // Uses the default "Financial Auditor" author

    // 2. Advanced: Modify the physical shape of the Note box
    // To modify the shape, we must query the worksheet's comments() collection.
    auto shape = wks.comments().shape("C2");
    
    // Make the comment box always visible
    shape.style().show();
    
    // Resize the comment box
    shape.style().setWidth(300);
    shape.style().setHeight(150);
```

---

## 4. Removing Comments

Removing comments can be done directly from the worksheet.

```cpp
    // 1. Remove a modern threaded comment and all its replies
    wks.deleteComment("A3");
    
    // 2. Remove a traditional note
    wks.deleteNote("C2");

    doc.save();
    doc.close();
    
    return 0;
}
```
