#ifndef OPENXLSX_XLDRAWING_HPP
#define OPENXLSX_XLDRAWING_HPP

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <ostream>    // std::basic_ostream

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
// #include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"

#include <memory>

namespace OpenXLSX
{
    class XLRelationships;

    // <v:fill o:detectmouseclick="t" type="solid" color2="#00003f"/>
    // <v:shadow on="t" obscured="t" color="black"/>
    // <v:stroke color="#3465a4" startarrow="block" startarrowwidth="medium" startarrowlength="medium" joinstyle="round" endcap="flat"/>
    // <v:path o:connecttype="none"/>
    // <v:textbox style="mso-direction-alt:auto;mso-fit-shape-to-text:t;">
    // 	<div style="text-align:left;"/>
    // </v:textbox>

    extern const std::string ShapeNodeName;        // = "v:shape"
    extern const std::string ShapeTypeNodeName;    // = "v:shapetype"

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
    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
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

        /**
         * @brief
         */
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
         * @brief Getter functions
         */
        std::string objectType() const;    // attribute ObjectType, value "Note"
        bool moveWithCells() const;    // element x:MoveWithCells - true = present or lowercase node_pcdata "true", false = not present or
                                       // lowercase node_pcdata "false"
        bool sizeWithCells() const;    // element x:SizeWithCells - logic as in MoveWithCells
        std::string
             anchor() const;    // element x:Anchor - Example node_pcdata: "3, 23, 0, 0, 4, 25, 3, 5" - no idea what any number means - TBD
        bool autoFill() const;                   // element x:AutoFill - logic as in MoveWithCells
        XLShapeTextVAlign textVAlign() const;    // element x:TextVAlign - Top, ???
        XLShapeTextHAlign textHAlign() const;    // element x:TextHAlign - Left, ???
        uint32_t          row() const;           // element x:Row, 0-indexed row of cell to which this shape is linked
        uint16_t          column() const;        // element x:Column, 0-indexed column of cell to which this shape is linked

        /**
         * @brief Setter functions
         */
        bool setObjectType(std::string newObjectType);
        bool setMoveWithCells(bool set = true);
        bool setSizeWithCells(bool set = true);
        bool setAnchor(std::string newAnchor);
        bool setAutoFill(bool set = true);
        bool setTextVAlign(XLShapeTextVAlign newTextVAlign);
        bool setTextHAlign(XLShapeTextHAlign newTextHAlign);
        bool setRow(uint32_t newRow);
        bool setColumn(uint16_t newColumn);

        // /**
        //  * @brief Return a string summary of the x:ClientData properties
        //  * @return string with info about the x:ClientData object
        //  */
        // std::string summary() const;

    private:                                                                // ---------- Private Member Variables ---------- //
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
        /**
         * @brief
         */
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
        int16_t attributeOrderIndex(std::string const& attributeName) const;
        /**
         * @brief XLShapeStyle internal generic getter & setter functions
         */
        XLShapeStyleAttribute getAttribute(std::string const& attributeName, std::string const& valIfNotFound = "") const;
        bool                  setAttribute(std::string const& attributeName, std::string const& attributeValue);
        // bool setAttribute(XLShapeStyleAttribute const& attribute);

    public:
        /**
         * @brief XLShapeStyle getter functions
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
         * @brief XLShapeStyle setter functions
         */
        bool setPosition(std::string newPosition);
        bool setMarginLeft(uint16_t newMarginLeft);
        bool setMarginTop(uint16_t newMarginTop);
        bool setWidth(uint16_t newWidth);
        bool setHeight(uint16_t newHeight);
        bool setMsoWrapStyle(std::string newMsoWrapStyle);
        bool setVTextAnchor(std::string newVTextAnchor);
        bool hide();    // set visibility:hidden
        bool show();    // set visibility:visible

        bool setRaw(std::string newStyle)
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
    public:                           // ---------- Public Member Functions ---------- //
        /**
         * @brief
         */
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

        /**
         * @brief
         */
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
         * @brief Getter functions
         */
        std::string  shapeId() const;        // v:shape attribute id - shape_# - can't be set by the user
        std::string  fillColor() const;      // v:shape attribute fillcolor, #<3 byte hex code>, e.g. #ffffc0
        bool         stroked() const;        // v:shape attribute stroked "t" ("f"?)
        std::string  type() const;           // v:shape attribute type, link to v:shapetype attribute id
        bool         allowInCell() const;    // v:shape attribute o:allowincell "f"
        XLShapeStyle style();                // v:shape attribute style, but constructed from the XMLAttribute

        // XLShapeShadow& shadow();          // v:shape subnode v:shadow
        // XLShapeFill& fill();              // v:shape subnode v:fill
        // XLShapeStroke& stroke();          // v:shape subnode v:stroke
        // XLShapePath& path();              // v:shape subnode v:path
        // XLShapeTextbox& textbox();        // v:shape subnode v:textbox
        XLShapeClientData clientData();    // v:shape subnode x:ClientData

        /**
         * @brief Setter functions
         * @param value that shall be set
         * @return true for success, false for failure
         */
        // NOTE: setShapeId is not available because shape id is managed by the parent class in createShape
        bool setFillColor(std::string const& newFillColor);
        bool setStroked(bool set);
        bool setType(std::string const& newType);
        bool setAllowInCell(bool set);
        bool setStyle(std::string const& newStyle);
        bool setStyle(XLShapeStyle const& newStyle);

    private:                                                           // ---------- Private Member Variables ---------- //
        mutable XMLNode                                   m_shapeNode; /**< An XMLNode object with the v:shape item */
        inline static const std::vector<std::string_view> m_nodeOrder =
            {"v:shadow", "v:fill", "v:stroke", "v:path", "v:textbox", "x:ClientData"};
    };

    /**
     * @brief An encapsulation of a drawing item (e.g. an image)
     */
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
        XLDrawing(XLXmlData* xmlData);

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
         */
        void addImage(const std::string& rId,
                      const std::string& name,
                      const std::string& description,
                      uint32_t           row,
                      uint32_t           col,
                      uint32_t           width,
                      uint32_t           height);

        /**
         * @brief Add a chart to the drawing
         * @param rId the relationship ID of the chart file
         * @param name the name of the chart
         * @param row the row where the chart should be placed (0-indexed)
         * @param col the column where the chart should be placed (0-indexed)
         * @param width the width of the chart in pixels
         * @param height the height of the chart in pixels
         */
        void addChartAnchor(const std::string& rId,
                            std::string_view   name,
                            uint32_t           row,
                            uint32_t           col,
                            uint32_t           width,
                            uint32_t           height);

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
        void addScaledImage(const std::string& rId,
                            const std::string& name,
                            const std::string& description,
                            const std::string& data,
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
        XLVmlDrawing(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLVmlDrawing(const XLVmlDrawing& other) = default;

        /**
         * @brief
         * @param other
         */
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
         * @brief
         * @param other
         * @return
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
        XMLNode shapeNode(std::string const& cellRef) const;

        uint32_t shapeCount() const;

        XLShape shape(uint32_t index) const;

        bool deleteShape(uint32_t index);
        bool deleteShape(std::string const& cellRef);

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
