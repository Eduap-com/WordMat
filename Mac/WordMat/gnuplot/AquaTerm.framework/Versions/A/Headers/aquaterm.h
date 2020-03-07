//
//  aquaterm.h
//  AquaTerm
//
//  Created by Per Persson on Sat Jul 12 2003.
//  Copyright (c) 2003-2004 AquaTerm. 
//

#include <stdint.h>

#define AQT_EVENTBUF_SIZE 128

/*" Constants that specify linecap styles. "*/
extern const int32_t AQTButtLineCapStyle;
extern const int32_t AQTRoundLineCapStyle;
extern const int32_t AQTSquareLineCapStyle;

/*" Constants that specify horizontal alignment for labels. "*/
extern const int32_t AQTAlignLeft;
extern const int32_t AQTAlignCenter;
extern const int32_t AQTAlignRight;
/*" Constants that specify vertical alignment for labels. "*/
extern const int32_t AQTAlignMiddle;
extern const int32_t AQTAlignBaseline;
extern const int32_t AQTAlignBottom;
extern const int32_t AQTAlignTop;

/*" Class initialization etc."*/
int32_t aqtInit(void);
void aqtTerminate(void);
/* The event handler callback functionality should be used with caution, it may 
   not be safe to use in all circumstances. It is certainly _not_ threadsafe. 
   If in doubt, use aqtWaitNextEvent() instead. */
void aqtSetEventHandler(void (*func)(int32_t ref, const char *event));

/*" Control operations "*/
void aqtOpenPlot(int32_t refNum);
int32_t aqtSelectPlot(int32_t refNum);
void aqtSetPlotSize(float width, float height);
void aqtSetPlotTitle(const char *title);
void aqtRenderPlot(void);
void aqtClearPlot(void);
void aqtClosePlot(void);

/*" Event handling "*/
void aqtSetAcceptingEvents(int32_t flag);
int32_t aqtGetLastEvent(char *buffer);
int32_t aqtWaitNextEvent(char *buffer);

/*" Plotting related commands "*/

/*" Clip rect, applies to all objects "*/
void aqtSetClipRect(float originX, float originY, float width, float height);
void aqtSetDefaultClipRect(void);

/*" Colormap (utility  "*/
int32_t aqtColormapSize(void);
void aqtSetColormapEntryRGBA(int32_t entryIndex, float r, float g, float b, float a);
void aqtGetColormapEntryRGBA(int32_t entryIndex, float *r, float *g, float *b, float *a);
void aqtSetColormapEntry(int32_t entryIndex, float r, float g, float b);
void aqtGetColormapEntry(int32_t entryIndex, float *r, float *g, float *b);
void aqtTakeColorFromColormapEntry(int32_t index);
void aqtTakeBackgroundColorFromColormapEntry(int32_t index);

/*" Color handling "*/
void aqtSetColorRGBA(float r, float g, float b, float a);
void aqtSetBackgroundColorRGBA(float r, float g, float b, float a);
void aqtGetColorRGBA(float *r, float *g, float *b, float *a);
void aqtGetBackgroundColorRGBA(float *r, float *g, float *b, float *a);
void aqtSetColor(float r, float g, float b);
void aqtSetBackgroundColor(float r, float g, float b);
void aqtGetColor(float *r, float *g, float *b);
void aqtGetBackgroundColor(float *r, float *g, float *b);

/*" Text handling "*/
void aqtSetFontname(const char *newFontname);
void aqtSetFontsize(float newFontsize);
void aqtAddLabel(const char *text, float x, float y, float angle, int32_t align);
void aqtAddShearedLabel(const char *text, float x, float y, float angle, float shearAngle, int32_t align);

/*" Line handling "*/
void aqtSetLinewidth(float newLinewidth);
void aqtSetLinestylePattern(float *newPattern, int32_t newCount, float newPhase);
void aqtSetLinestyleSolid(void);
void aqtSetLineCapStyle(int32_t capStyle);
void aqtMoveTo(float x, float y);
void aqtAddLineTo(float x, float y);
void aqtAddPolyline(float *x, float *y, int32_t pointCount);

/*" Rect and polygon handling"*/
void aqtMoveToVertex(float x, float y);
void aqtAddEdgeToVertex(float x, float y);
void aqtAddPolygon(float *x, float *y, int32_t pointCount);
void aqtAddFilledRect(float originX, float originY, float width, float height);
void aqtEraseRect(float originX, float originY, float width, float height);

/*" Image handling "*/
void aqtSetImageTransform(float m11, float m12, float m21, float m22, float tX, float tY);
void aqtResetImageTransform(void);
void aqtAddImageWithBitmap(const void *bitmap, int32_t pixWide, int32_t pixHigh, float destX, float destY, float destWidth, float destHeight);
void aqtAddTransformedImageWithBitmap(const void *bitmap, int32_t pixWide, int32_t pixHigh, float clipX, float clipY, float clipWidth, float clipHeight);

