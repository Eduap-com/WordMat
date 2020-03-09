//
//  AQTAdapter.h
//  AquaTerm
//
//  Created by Per Persson on Sat Jul 12 2003.
//  Copyright (c) 2003-2004 AquaTerm.
//

#import <Foundation/NSString.h>
#import <Foundation/NSGeometry.h>

/*" Constants that specify linecap styles. "*/
extern const int32_t AQTButtLineCapStyle;
extern const int32_t AQTRoundLineCapStyle;
extern const int32_t AQTSquareLineCapStyle;

/*" Constants that specify horizontal and vertical alignment for labels. See #addLabel:atPoint:angle:align: for definitions and use."*/
extern const int32_t AQTAlignLeft;
extern const int32_t AQTAlignCenter;
extern const int32_t AQTAlignRight;
/* Constants that specify vertical alignment for labels. */
extern const int32_t AQTAlignMiddle;
extern const int32_t AQTAlignBaseline;
extern const int32_t AQTAlignBottom;
extern const int32_t AQTAlignTop;

@class AQTPlotBuilder, AQTClientManager;
@interface AQTAdapter : NSObject
{
   /*" All instance variables are private. "*/
   AQTClientManager *_clientManager;
   AQTPlotBuilder *_selectedBuilder;
   id _aqtReserved1;
   id _aqtReserved2;
}

/*" Class initialization etc."*/
- (id)init;
- (id)initWithServer:(id)localServer;
- (void)setErrorHandler:(void (*)(NSString *msg))fPtr;
- (void)setEventHandler:(void (*)(int32_t index, NSString *event))fPtr;

  /*" Control operations "*/
- (void)openPlotWithIndex:(int32_t)refNum; 
- (BOOL)selectPlotWithIndex:(int32_t)refNum;
- (void)setPlotSize:(NSSize)canvasSize;
- (void)setPlotTitle:(NSString *)title;
- (void)renderPlot;
- (void)clearPlot;
- (void)closePlot;

  /*" Event handling "*/
- (void)setAcceptingEvents:(BOOL)flag;
- (NSString *)lastEvent;
- (NSString *)waitNextEvent; 

/*" Plotting related commands "*/

/*" Clip rect, applies to all objects "*/
- (void)setClipRect:(NSRect)clip;
- (void)setDefaultClipRect;

/*" Colormap (utility) "*/
- (int32_t)colormapSize;
- (void)setColormapEntry:(int32_t)entryIndex red:(float)r green:(float)g blue:(float)b alpha:(float)a;
- (void)getColormapEntry:(int32_t)entryIndex red:(float *)r green:(float *)g blue:(float *)b alpha:(float *)a;
- (void)setColormapEntry:(int32_t)entryIndex red:(float)r green:(float)g blue:(float)b;
- (void)getColormapEntry:(int32_t)entryIndex red:(float *)r green:(float *)g blue:(float *)b;
- (void)takeColorFromColormapEntry:(int32_t)index;
- (void)takeBackgroundColorFromColormapEntry:(int32_t)index;

  /*" Color handling "*/
- (void)setColorRed:(float)r green:(float)g blue:(float)b alpha:(float)a;
- (void)setBackgroundColorRed:(float)r green:(float)g blue:(float)b alpha:(float)a;
- (void)getColorRed:(float *)r green:(float *)g blue:(float *)b alpha:(float *)a;
- (void)getBackgroundColorRed:(float *)r green:(float *)g blue:(float *)b alpha:(float *)a;
- (void)setColorRed:(float)r green:(float)g blue:(float)b;
- (void)setBackgroundColorRed:(float)r green:(float)g blue:(float)b;
- (void)getColorRed:(float *)r green:(float *)g blue:(float *)b;
- (void)getBackgroundColorRed:(float *)r green:(float *)g blue:(float *)b;

  /*" Text handling "*/
- (void)setFontname:(NSString *)newFontname;
- (void)setFontsize:(float)newFontsize;
- (void)addLabel:(id)text atPoint:(NSPoint)pos;
- (void)addLabel:(id)text atPoint:(NSPoint)pos angle:(float)angle align:(int32_t)just;
- (void)addLabel:(id)text atPoint:(NSPoint)pos angle:(float)angle shearAngle:(float)shearAngle align:(int32_t)just;

  /*" Line handling "*/
- (void)setLinewidth:(float)newLinewidth;
- (void)setLinestylePattern:(float *)newPattern count:(int32_t)newCount phase:(float)newPhase;
- (void)setLinestyleSolid;
- (void)setLineCapStyle:(int32_t)capStyle;
- (void)moveToPoint:(NSPoint)point;  
- (void)addLineToPoint:(NSPoint)point; 
- (void)addPolylineWithPoints:(NSPoint *)points pointCount:(int32_t)pc;

  /*" Rect and polygon handling"*/
- (void)moveToVertexPoint:(NSPoint)point;
- (void)addEdgeToVertexPoint:(NSPoint)point; 
- (void)addPolygonWithVertexPoints:(NSPoint *)points pointCount:(int32_t)pc;
- (void)addFilledRect:(NSRect)aRect;
- (void)eraseRect:(NSRect)aRect;

  /*" Image handling "*/
- (void)setImageTransformM11:(float)m11 m12:(float)m12 m21:(float)m21 m22:(float)m22 tX:(float)tX tY:(float)tY;
- (void)resetImageTransform;
- (void)addImageWithBitmap:(const void *)bitmap size:(NSSize)bitmapSize bounds:(NSRect)destBounds; 
- (void)addTransformedImageWithBitmap:(const void *)bitmap size:(NSSize)bitmapSize clipRect:(NSRect)destBounds;
- (void)addTransformedImageWithBitmap:(const void *)bitmap size:(NSSize)bitmapSize;

  /*"Private methods"*/
- (void)timingTestWithTag:(uint32_t)tag;
@end
