#ifndef OPENXLSX_XLWORKSHEETIMPL_HPP
#define OPENXLSX_XLWORKSHEETIMPL_HPP

#include "XLRelationships.hpp"
#include "XLMergeCells.hpp"
#include "XLDataValidation.hpp"
#include "XLDrawing.hpp"
#include "XLComments.hpp"
#include "XLThreadedComments.hpp"
#include "XLTables.hpp"
#include "XLSlicerCollection.hpp"

namespace OpenXLSX {
    struct XLWorksheetImpl {
        XLRelationships                                    m_relationships{};
        XLMergeCells                                       m_merges{};
        XLDataValidations                                  m_dataValidations{};
        XLDrawing                                          m_drawing{};
        XLVmlDrawing                                       m_vmlDrawing{};
        XLComments                                         m_comments{};
        XLThreadedComments                                 m_threadedComments{};
        XLTableCollection                                  m_tables{};
        XLSlicerCollection                                 m_slicers{};
    };
}

#endif // OPENXLSX_XLWORKSHEETIMPL_HPP
