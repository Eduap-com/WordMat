# Shared

The share folder holds files which are uses in both the Windows and Mac version. The two installers grab the files from here.

Actually many files are in both versions, but must be copied from Windows to Mac, when they are changed and require a bit of tweaking.
Mainly because special characters can mess up in many files when they are opened on a Mac.

## Fonts
This folder holds math-fonts which are installed into Word.
The fonts are used to create documents that are similar to Latex-documents

## Maxima-files
Files that are added to the original Maxima installation or changes

## Translations
Excel sheet to translate WordMat

## WordDocs
Different Worddocuments used in WordMat
Word documents that contain code must be kept separately between Windows and Mac

## WordMat-Ribbon
The ribbon is in the WordMat.dotm file. This folder folds icons used in the ribbon in case they mess up.

## Testing
Files used to test WordMat before a new release

## GeoGebra math apps bundle
Javascript bundle to run GeoGebra in a browser
Version 5-0-694 used win v. 1.24
Version 5-0-723 used win v. 1.25
Version 5-0-791 used win v. 1.26
Version 5-0-805 used win v. 1.27

Get new version here:
https://geogebra.github.io/docs/reference/en/GeoGebra_Apps_Embedding/


###geogebra-math-apps\GeoGebra\HTML5\5.0\GeoGebra.html changed:

PerspectivePopup disabled. This is important, but you may need more:
In GeoGebra.html line 809:

.perspectivePopup,
.appPickerPopup {
	display: none !important;
}



Add to function loadApp() at line 423 (after updateAppletParams)

		// Suppress restore dialog for locally cached unsaved work.
		try { localStorage.clear(); } catch(e) {}

		// Suppress "leave without saving" beforeunload dialog.
		try {
			var _origAddEL = window.addEventListener;
			window.addEventListener = function(type, fn, opts) {
				if (type === 'beforeunload') return;
				return _origAddEL.call(this, type, fn, opts);
			};
			Object.defineProperty(window, 'onbeforeunload', { set: function(){}, get: function(){ return null; }, configurable: true });
		} catch(e) {}

		// Keep the on-screen math keyboard hidden at startup.
		// It is enabled again after the first input/text interaction.
		try {
			document.body.classList.add("keyboard-initial-hidden");
			var isTextInputTarget = function(target) {
				return target && target.closest && target.closest("input, textarea, [contenteditable='true'], .AutoCompleteTextFieldW, [role='textbox']");
			};
			var setInitialHeightOverride = function(el, height) {
				if (!el || !el.style) {
					return;
				}
				if (!el.hasAttribute("data-initial-kb-height")) {
					el.setAttribute("data-initial-kb-height", "1");
					el.setAttribute("data-initial-kb-prev-height", el.style.height || "");
				}
				el.style.height = height + "px";
			};
			var applyInitialLayoutOverride = function() {
				if (!document.body.classList.contains("keyboard-initial-hidden")) {
					return;
				}
				var panels = document.querySelectorAll(".splitPanelWrapper, .gwt-SplitLayoutPanel");
				for (var i = 0; i < panels.length; i++) {
					var panel = panels[i];
					var top = panel.getBoundingClientRect().top;
					var height = Math.max(0, window.innerHeight - top);
					setInitialHeightOverride(panel, height);
					var nested = panel.querySelectorAll(".dockPanelParent, .ggbdockpanelhack, .EuclidianPanel, .EuclidianPanel canvas, .gwt-SplitLayoutPanel-HDragger");
					for (var j = 0; j < nested.length; j++) {
						setInitialHeightOverride(nested[j], height);
					}
				}
			};
			var clearInitialLayoutOverride = function() {
				var panels = document.querySelectorAll("[data-initial-kb-height='1']");
				for (var i = 0; i < panels.length; i++) {
					panels[i].style.height = panels[i].getAttribute("data-initial-kb-prev-height") || "";
					panels[i].removeAttribute("data-initial-kb-prev-height");
					panels[i].removeAttribute("data-initial-kb-height");
				}
			};
			var releaseInitialKeyboardHide = function(evt) {
				var target = evt && evt.target;
				if (!target || !target.closest) {
					return;
				}
				var inputTarget = isTextInputTarget(target);
				if (inputTarget) {
					document.body.classList.remove("keyboard-initial-hidden");
					document.removeEventListener("pointerdown", releaseInitialKeyboardHide, true);
					window.removeEventListener("resize", applyInitialLayoutOverride);
					clearInitialLayoutOverride();
					window.dispatchEvent(new Event("resize"));
				}
			};
			document.addEventListener("pointerdown", releaseInitialKeyboardHide, true);
			window.addEventListener("resize", applyInitialLayoutOverride);
			applyInitialLayoutOverride();
			window.setTimeout(function() {
				applyInitialLayoutOverride();
			}, 120);
			window.setTimeout(function() {
				applyInitialLayoutOverride();
			}, 500);
		} catch (e) {}
