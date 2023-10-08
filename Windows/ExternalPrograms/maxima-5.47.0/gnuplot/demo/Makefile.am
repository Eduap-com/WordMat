## Process this file with automake to produce Makefile.in -*-Makefile-*-
AUTOMAKE_OPTIONS = foreign

CLEANFILES = binary1 binary2 binary3 defaults.ini equipo2.tmp field2xy.tmp \
soundfit.par temp.set fontfile.ps fontfile_latex.ps epslatex-inc.eps \
epslatex-inc.pdf epslatex.aux epslatex.dvi epslatex.log epslatex.pdf \
epslatex.ps epslatex.tex random.tmp stringvar.tmp fit.log fitmulti.dat

BINARY_FILES = binary1 binary2 binary3

SUBDIRS = $(PLUGIN_DIR)
DIST_SUBDIRS = plugin

DEMO = all.dem

GNUPLOT = gnuplot

all: $(BINARY_FILES)

$(BINARY_FILES): ../src/bf_test
	@echo Creating binary data files
	@../src/bf_test

check-prepare:

check-local: check-noninteractive

check-interactive: check-prepare $(BINARY_FILES)
	@if test -z "$(GNUTERM)" ; then \
	  ( bdir=`pwd` ; PATH=$$bdir/../src:$$PATH \
	    FIT_LOG="" \
	    GNUPLOT_DRIVER_DIR=$$bdir/../src \
	    GNUPLOT_LIB=$(srcdir) $(GNUPLOT) $(DEMO) ); \
	else \
	  ( bdir=`pwd` ; PATH=$$bdir/../src:$$PATH \
	    FIT_LOG="" \
	    GNUPLOT_DRIVER_DIR=$$bdir/../src \
	    GNUTERM="$(GNUTERM)" \
	    GNUPLOT_LIB=$(srcdir) $(GNUPLOT) $(DEMO) ); \
	fi

check-noninteractive: check-prepare $(BINARY_FILES)
	@if test -z "$(GNUTERM)" ; then \
	  ( bdir=`pwd` ; PATH=$$bdir/../src:$$PATH \
	    FIT_LOG="" \
	    GNUPLOT_DRIVER_DIR=$$bdir/../src \
	    GNUPLOT_LIB=$(srcdir) $(GNUPLOT) $(DEMO) </dev/null ); \
	else \
	  ( bdir=`pwd` ; PATH=$$bdir/../src:$$PATH\
	    FIT_LOG="" \
	    GNUPLOT_DRIVER_DIR=$$bdir/../src \
	    GNUTERM="$(GNUTERM)" \
	    GNUPLOT_LIB=$(srcdir) $(GNUPLOT) $(DEMO) </dev/null ); \
	fi

epslatex:
	@if test -z "$(GNUTERM)" ; then \
	  ( bdir=`pwd` ; PATH=$$bdir/../src:$$PATH \
	    GNUPLOT_DRIVER_DIR=$$bdir/../src \
	    GNUPLOT_LIB=$(srcdir) $(GNUPLOT) epslatex.dem ); \
	else \
	  ( bdir=`pwd` ; PATH=$$bdir/../src:$$PATH \
	    GNUPLOT_DRIVER_DIR=$$bdir/../src \
	    GNUTERM="$(GNUTERM)" \
	    GNUPLOT_LIB=$(srcdir) $(GNUPLOT) epslatex.dem ); \
	fi
	latex epslatex
	dvips epslatex

##m4-files-begin
EXTRA_DIST = Makefile.am.in lena-keypoints.bin scatter2.bin sine.bin \
using.bin gnuplot.cfg world.cor 1.dat 2.dat 3.dat arrowstyle.dat \
asciimat.dat battery.dat big_peak.dat candlesticks.dat cities.dat \
clip14in.dat ctg-y2.dat delaunay-edges.dat dodecahedron.dat ellipses.dat \
empty-circles.dat energy_circles.dat finance.dat fit3.dat glass.dat \
hemisphr.dat histerror.dat histopt.dat icosahedron.dat immigration.dat \
klein.dat lcdemo.dat mask_pm3d.dat moli3.dat optimize.dat \
orbital_elements.dat rgb_variable.dat scatter2.dat silver.dat \
soundvel.dat srl.dat steps.dat table.dat timedat.dat triangle.dat \
truncated_cube.dat using.dat whale.dat world.dat WorldEnergy.dat \
3Dboxes.dem airfoil.dem all.dem animate2.dem animate.dem approximate.dem \
armillary.dem array.dem arrows.dem arrowstyle.dem autoscale.dem \
azimuth.dem barchart_art.dem binary.dem bins.dem bivariat.dem \
bolditalic.dem borders.dem boxclusters.dem boxes3d.dem boxplot.dem \
break_continue.dem callargs.dem candlesticks.dem cerf.dem charset.dem \
circles.dem cities.dem clipobject.dem clip_radial.dem colorscheme.dem \
colorwheel.dem columnhead.dem complex_trig.dem contours.dem controls.dem \
custom_contours.dem custom_key.dem dashcolor.dem dashtypes.dem \
datastrings.dem dgrid3d.dem discrete.dem electron.dem ellipse.dem \
ellipses_style.dem enhancedtext.dem enhanced_utf8.dem epi_data.dem \
epslatex.dem errorbars.dem fenceplot.dem fillbetween.dem fillcrvs.dem \
fillstyle.dem finance.dem fit.dem fitmulti.dem fontfile.dem \
fontfile_latex.dem gantt.dem heatmap_points.dem heatmaps.dem \
hidden2.dem hidden_compare.dem hidden.dem histerror.dem histograms2.dem \
histograms.dem hypertext.dem image2.dem image.dem imageNaN.dem \
int64.dem isosurface.dem iterate.dem jitter.dem kdensity2d.dem key.dem \
keyentry.dem label_stacked_histograms.dem layout.dem lines_arrows.dem \
linked_autoscale.dem linkedaxes.dem logscale_clipping.dem macros.dem \
map_projection.dem margins.dem matrix_every.dem matrix_index.dem \
mgr.dem molecule.dem monotonic_spline.dem mouselab_1.dem mouselab_2.dem \
mouselabels.dem mousevariables.dem multiaxis.dem multipalette.dem \
multiplt.dem named_var.dem nokey.dem nonlinear1.dem nonlinear2.dem \
nonlinear3.dem nonlinear4.dem nonlinear5.dem nonlinear6.dem orbits.dem \
overflow.dem palette+alpha.dem parallel.dem param.dem piecewise.dem \
pixmap.dem pm3d_clip.dem pm3dcolors.dem pm3d.dem pm3dgamma.dem \
pm3d_lighting.dem pointsize.dem polar.dem polar_quadrants.dem \
poldat.dem polygons.dem prob2.dem probably_tux.dem prob.dem \
projection.dem pt_variable.dem rainbow.dem random.dem rectangle.dem \
rgba_lines.dem rgbalpha.dem rgb_variable.dem rotate_labels.dem \
rugplot.dem running_avg.dem sampling2.dem sampling.dem scatter.dem \
short_vector.dem simple.dem singulr.dem smooth.dem smooth_splines.dem \
smooth_zsort.dem solar_params.dem solar_path.dem special_chars.dem \
spiderplot.dem spline.dem stats.dem steps.dem stringvar.dem surface1.dem \
surface2.dem textbox.dem textcolor.dem textrotate.dem tics.dem timedat.dem \
transparent.dem transparent_solids.dem ttics.dem unicode.dem using.dem \
utf8.dem varcolor.dem vector.dem violinplot.dem volatile.dem voxel.dem \
vplot.dem walls.dem week_date.dem world2.dem world.dem zerror.dem \
demo.edf density.fnc hexa.fnc line.fnc reflect.fnc gen-random.inc stat.inc \
GM1_sugar.pdb aries.png bldg.png gradient.png icon64x64.png GM1_bonds.r3d \
blutux.rgb lena.rgb sound.par sound2.par start.par gnuplot.rot gnu-valley \
random-points gpdemos.tcl html 
##m4-files-end

Makefile.am: Makefile.am.in
	rm -f $@ $@t
	sed -n '1,/^##m4-files-begin/p' Makefile.am.in > $@t
	echo EXTRA_DIST = Makefile.am.in *.bin *.cfg *.cor *.dat *.dem *.edf \
	  *.fnc *.inc *.pdb *.png *.r3d *.rgb sound.par sound2.par \
	  start.par *.rot gnu-valley random-points gpdemos.tcl html | fmt | \
	  (tr '\012' @; echo) | sed 's/@$$/%/;s/@/ \\@/g' | tr @% '\012 ' \
	  >> $@t
	sed -n '/^##m4-files-end/,$$p' $< >> $@t
	echo if BUILD_PLUGIN >> $@t
	echo    PLUGIN_DIR = plugin >> $@t
	echo endif >> $@t
	chmod og-w $@t
	mv $@t $@

if BUILD_PLUGIN
PLUGIN_DIR = plugin
endif
