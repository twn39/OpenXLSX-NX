#ifndef OPENXLSX_XLDRAWING_HPP
#define OPENXLSX_XLDRAWING_HPP

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <gsl/pointers>
#include <ostream>    // std::basic_ostream
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
// #include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"
#include "XLConstants.hpp"
#include "XLImageOptions.hpp"

#include <memory>

namespace OpenXLSX
{
    class XLRelationships;

    extern const std::string_view ShapeNodeName;        // = "v:shape"
    extern const std::string_view ShapeTypeNodeName;    // = "v:shapetype"

    // NOTE: numerical values of XLShapeTextVAlign and XLShapeTextHAlign are shared with the same alignments from XLAlignmentStyle
    // (XLStyles.hpp)
    enum class XLShapeTextVAlign : uint8_t {
        Center  = 3,     // value="center",           both
        Top     = 4,     // value="top",              vertical only
        Bottom  = 5,     // value="bottom",           vertical only
        Invalid = 255    // all other values
    };
    constexpr const XLShapeTextVAlign XLDefaultShapeTextVAlign = XLShapeTextVAlign::Top;

    enum class XLShapeTextHAlign : uint8_t {
        Left    = 1,     // value="left",             horizontal only
        Right   = 2,     // value="right",            horizontal only
        Center  = 3,     // value="center",           both
        Invalid = 255    // all other values
    };
    constexpr const XLShapeTextHAlign XLDefaultShapeTextHAlign = XLShapeTextHAlign::Left;

    /**
     * @brief An encapsulation of a shape client data element x:ClientData
     */
    class OPENXLSX_EXPORT XLShapeClientData
    {
    public:
XLShapeClientData();

        /**
         * @brief Constructor. New items should only be created through an XLShape object.
         * @param node An XMLNode object with the x:ClientData XMLNode. If no input is provided, a null node is used.
         */
        explicit XLShapeClientData(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLShapeClientData(const XLShapeClientData& other) = default;

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLShapeClientData(XLShapeClientData&& other) noexcept = default;

        ~XLShapeClientData() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLShapeClientData& operator=(const XLShapeClientData& other) = default;

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLShapeClientData& operator=(XLShapeClientData&& other) noexcept = default;

        /**
         * @brief Component Data Accessors
         */
        std::string objectType() const;    // attribute ObjectType, value "Note"
        bool moveWithCells() const;    // element x:MoveWithCells - true = present or lowercase node_pcdata "true", false = not present or
                                       // lowercase node_pcdata "false"
        bool sizeWithCells() const;    // element x:SizeWithCells - logic as in MoveWithCells
        std::string
             anchor() const;    // element x:Anchor - e.g. "3, 23, 0, 0, 4, 25, 3, 5"
        bool autoFill() const;                   // element x:AutoFill - logic as in MoveWithCells
        XLShapeTextVAlign textVAlign() const;    // element x:TextVAlign - Top, ???
        XLShapeTextHAlign textHAlign() const;    // element x:TextHAlign - Left, ???
        uint32_t          row() const;           // element x:Row, 0-indexed row of cell to which this shape is linked
        uint16_t          column() const;        // element x:Column, 0-indexed column of cell to which this shape is linked

        /**
         * @brief Component Data Mutators
         */
        bool setObjectType(std::string_view newObjectType);
        bool setMoveWithCells(bool set = true);
        bool setSizeWithCells(bool set = true);
        bool setAnchor(std::string_view newAnchor);
        bool setAutoFill(bool set = true);
        bool setTextVAlign(XLShapeTextVAlign newTextVAlign);
        bool setTextHAlign(XLShapeTextHAlign newTextHAlign);
        bool setRow(uint32_t newRow);
        bool setColumn(uint16_t newColumn);

    private:
mutable XMLNode                                   m_clientDataNode; /**< An XMLNode object with the x:ClientData item */
        inline static const std::vector<std::string_view> m_nodeOrder =
            {"x:MoveWithCells", "x:SizeWithCells", "x:Anchor", "x:AutoFill", "x:TextVAlign", "x:TextHAlign", "x:Row", "x:Column"};
    };

    struct XLShapeStyleAttribute
    {
        std::string name;
        std::string value;
    };

    class OPENXLSX_EXPORT XLShapeStyle
    {
    public:
        XLShapeStyle();

        /**
         * @brief Constructor. Init XLShapeStyle properties from styleAttribute
         * @param styleAttribute a string with the value of the style attribute of a v:shape element
         */
        explicit XLShapeStyle(const std::string& styleAttribute);

        /**
         * @brief Constructor. Init XLShapeStyle properties from styleAttribute and link to the attribute so that setter functions directly
         * modify it
         * @param styleAttribute an XMLAttribute constructed with the style attribute of a v:shape element
         */
        explicit XLShapeStyle(const XMLAttribute& styleAttribute);

    private:
        /**
         * @brief get index of an attribute name within m_nodeOrder
         * @return index of attribute in m_nodeOrder
         * @return -1 if not found
         */
        int16_t attributeOrderIndex(std::string_view attributeName) const;
        /**
         * @brief XLShapeStyle internal generic getter & setter functions
         */
        XLShapeStyleAttribute getAttribute(std::string_view attributeName, std::string_view valIfNotFound = "") const;
        bool                  setAttribute(std::string_view attributeName, std::string_view attributeValue);

    public:
        /**
         * @brief Style Attribute Accessors
         */
        std::string position() const;
        uint16_t    marginLeft() const;
        uint16_t    marginTop() const;
        uint16_t    width() const;
        uint16_t    height() const;
        std::string msoWrapStyle() const;
        std::string vTextAnchor() const;
        bool        hidden() const;
        bool        visible() const;

        std::string raw() const { return m_style; }

        /**
         * @brief Style Attribute Mutators
         */
        bool setPosition(std::string_view newPosition);
        bool setMarginLeft(uint16_t newMarginLeft);
        bool setMarginTop(uint16_t newMarginTop);
        bool setWidth(uint16_t newWidth);
        bool setHeight(uint16_t newHeight);
        bool setMsoWrapStyle(std::string_view newMsoWrapStyle);
        bool setVTextAnchor(std::string_view newVTextAnchor);
        bool hide();    // set visibility:hidden
        bool show();    // set visibility:visible

        bool setRaw(std::string_view newStyle)
        {
            m_style = newStyle;
            return true;
        }

    private:
        mutable std::string  m_style;    // mutable so getter functions can update it from m_styleAttribute if the latter is not empty
        mutable XMLAttribute m_styleAttribute;
        inline static const std::vector<std::string_view> m_nodeOrder = {"position",
                                                                         "margin-left",
                                                                         "margin-top",
                                                                         "width",
                                                                         "height",
                                                                         "mso-wrap-style",
                                                                         "mso-fit-shape-to-text",    // Added for auto-size support
                                                                         "v-text-anchor",
                                                                         "visibility"};
    };

    class OPENXLSX_EXPORT XLShape
    {
        friend class XLVmlDrawing;    // for access to m_shapeNode in XLVmlDrawing::addShape
        friend class XLComments;      // added for access to m_shapeNode in XLComments::set
    public:
XLShape();

        /**
         * @brief Constructor. New items should only be created through an XLStyles object.
         * @param node An XMLNode object with the styles item. If no input is provided, a null node is used.
         */
        explicit XLShape(const XMLNode& node);

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLShape(const XLShape& other) = default;

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLShape(XLShape&& other) noexcept = default;

        ~XLShape() = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLShape& operator=(const XLShape& other);

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLShape& operator=(XLShape&& other) noexcept = default;

        /**
         * @brief Component Data Accessors
         */
        std::string  shapeId() const;        // v:shape attribute id - shape_# - can't be set by the user
        std::string  fillColor() const;      // v:shape attribute fillcolor, #<3 byte hex code>, e.g. #ffffc0
        bool         stroked() const;        // v:shape attribute stroked "t" ("f"?)
        std::string  type() const;           // v:shape attribute type, link to v:shapetype attribute id
        bool         allowInCell() const;    // v:shape attribute o:allowincell "f"
        XLShapeStyle style();                // v:shape attribute style, but constructed from the XMLAttribute

        XLShapeClientData clientData();    // v:shape subnode x:ClientData

        /**
         * @brief Setter functions
         * @param value that shall be set
         * @return true for success, false for failure
         */
        // NOTE: setShapeId is not available because shape id is managed by the parent class in createShape
        bool setFillColor(std::string_view newFillColor);
        bool setStroked(bool set);
        bool setType(std::string_view newType);
        bool setAllowInCell(bool set);
        bool setStyle(std::string_view newStyle);
        bool setStyle(XLShapeStyle const& newStyle);
    private:
mutable XMLNode                                   m_shapeNode; /**< An XMLNode object with the v:shape item */
        inline static const std::vector<std::string_view> m_nodeOrder =
            {"v:shadow", "v:fill", "v:stroke", "v:path", "v:textbox", "x:ClientData"};
    };

    /**
     * @brief An encapsulation of a drawing item (e.g. an image)
     */

    enum class XLVectorShapeType {
        Rectangle,
        Ellipse,
        Line,
        Triangle,
        RightTriangle,
        Arrow,
        Diamond,
        Parallelogram,
        Hexagon
    };

    struct XLVectorShapeOptions {
        XLVectorShapeType type = XLVectorShapeType::Rectangle;
        std::string name = "Shape 1";
        std::string text = "";
        std::string fillColor = "4286F4"; // ARGB, no #
        std::string lineColor = "000000";
        double lineWidth = 1.0;

        uint32_t width = 100;
        uint32_t height = 100;
        int32_t offsetX = 0;
        int32_t offsetY = 0;

        std::string macro = "";
    };
    class OPENXLSX_EXPORT XLDrawingItem
    {
    public:
        XLDrawingItem();
        explicit XLDrawingItem(const XMLNode& node);
        XLDrawingItem(const XLDrawingItem& other)                = default;
        XLDrawingItem(XLDrawingItem&& other) noexcept            = default;
        ~XLDrawingItem()                                         = default;
        XLDrawingItem& operator=(const XLDrawingItem& other)     = default;
        XLDrawingItem& operator=(XLDrawingItem&& other) noexcept = default;

        std::string name() const;
        std::string description() const;
        uint32_t    row() const;
        uint32_t    col() const;
        uint32_t    width() const;     // in pixels (converted from EMUs)
        uint32_t    height() const;    // in pixels (converted from EMUs)
        std::string relationshipId() const;

    private:
        mutable XMLNode m_anchorNode;
    };

    /**
     * @brief The XLDrawing class is the base class for worksheet drawings (images, charts, etc.)
     */
    class OPENXLSX_EXPORT XLDrawing : public XLXmlFile
    {
        friend class XLWorksheet;
        friend class XLDocument;

    public:
        /**
         * @brief Constructor
         */
        XLDrawing() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor.
         * @param xmlData the source XML of the drawings file
         */
        explicit XLDrawing(gsl::not_null<XLXmlData*> xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         */
        XLDrawing(const XLDrawing& other);

        /**
         * @brief Move constructor.
         */
        XLDrawing(XLDrawing&& other) noexcept;

        /**
         * @brief The destructor
         */
        ~XLDrawing() = default;

        /**
         * @brief Assignment operator
         */
        XLDrawing& operator=(const XLDrawing& other);

        /**
         * @brief Move assignment operator.
         */
        XLDrawing& operator=(XLDrawing&& other) noexcept;

        /**
         * @brief Add an image to the drawing
         * @param rId the relationship ID of the image file
         * @param name the name of the image
         * @param description the description of the image
         * @param row the row where the image should be placed (0-indexed)
         * @param col the column where the image should be placed (0-indexed)
         * @param width the width of the image in pixels
         * @param height the height of the image in pixels
         * @param options the formatting and positioning options
         */
        void addImage(std::string_view rId,
                      std::string_view name,
                      std::string_view description,
                      uint32_t           row,
                      uint32_t           col,
                      uint32_t           width,
                      uint32_t           height,
                      const XLImageOptions& options = XLImageOptions());

        /**
         * @brief Add a chart to the drawing
         * @param rId the relationship ID of the chart file
         * @param name the name of the chart
         * @param row the row where the chart should be placed (0-indexed)
         * @param col the column where the chart should be placed (0-indexed)
         * @param width the width of the chart in pixels
         * @param height the height of the chart in pixels
         */

        void addShape(uint32_t row, uint32_t col, const XLVectorShapeOptions& options = XLVectorShapeOptions());
        void addChartAnchor(std::string_view rId,
                            std::string_view   name,
                            uint32_t           row,
                            uint32_t           col,
                            uint32_t           width,
                            uint32_t           height);

        void addChartAnchor(std::string_view rId,
                               std::string_view   name,
                               uint32_t           row,
                               uint32_t           col,
                               XLDistance         width,
                               XLDistance         height);

        /**
         * @brief Add an image to the drawing, automatically maintaining aspect ratio
         * @param rId the relationship ID of the image file
         * @param name the name of the image
         * @param description the description of the image
         * @param data the binary data of the image (used to detect dimensions)
         * @param row the row where the image should be placed (0-indexed)
         * @param col the column where the image should be placed (0-indexed)
         * @param scalingFactor the factor to scale the image by (1.0 = original size)
         */
        void addScaledImage(std::string_view rId,
                            std::string_view name,
                            std::string_view description,
                            std::string_view data,
                            uint32_t           row,
                            uint32_t           col,
                            double             scalingFactor = 1.0);

        /**
         * @brief Get the drawing relationships
         * @return a reference to the XLRelationships object
         */
        XLRelationships& relationships();

        /**
         * @brief Get the number of images in the drawing
         * @return the number of images
         */
        uint32_t imageCount() const;

        /**
         * @brief Get an image item by index
         * @param index the index of the image
         * @return an XLDrawingItem object
         */
        XLDrawingItem image(uint32_t index) const;

    private:
        // Helper to add namespaces if missing
        void initXml();

        std::unique_ptr<XLRelationships> m_relationships; /**< class handling the drawing relationships */
    };

    /**
     * @brief The XLVmlDrawing class is the base class for worksheet comments
     */
    class OPENXLSX_EXPORT XLVmlDrawing : public XLXmlFile
    {
        friend class XLWorksheet;    // for access to XLXmlFile::getXmlPath
        friend class XLComments;     // for access to firstShapeNode
    public:
        /**
         * @brief Constructor
         */
        XLVmlDrawing() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor.
         * @param xmlData the source XML of the comments file
         */
        explicit XLVmlDrawing(gsl::not_null<XLXmlData*> xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLVmlDrawing(const XLVmlDrawing& other) = default;

        XLVmlDrawing(XLVmlDrawing&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLVmlDrawing() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLVmlDrawing& operator=(const XLVmlDrawing&) = default;

        /**
         */
        XLVmlDrawing& operator=(XLVmlDrawing&& other) noexcept = default;

    private:    // helper functions with repeating code
        XMLNode firstShapeNode() const;
        XMLNode lastShapeNode() const;
        XMLNode shapeNode(uint32_t index) const;

    public:
        /**
         * @brief Get the shape XML node that is associated with the cell indicated by cellRef
         * @param cellRef the reference to the cell for which a shape shall be found
         * @return the XMLNode that contains the desired shape, or an empty XMLNode if not found
         */
        XMLNode shapeNode(std::string_view cellRef) const;

        uint32_t shapeCount() const;

        XLShape shape(uint32_t index) const;

        bool deleteShape(uint32_t index);
        bool deleteShape(std::string_view cellRef);

        XLShape createShape(const XLShape& shapeTemplate = XLShape());

        /**
         * @brief Print the XML contents of this XLVmlDrawing instance using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:
        uint32_t    m_shapeCount{0};
        uint32_t    m_lastAssignedShapeId{0};
        std::string m_defaultShapeTypeId{};
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLDRAWING_HPP
