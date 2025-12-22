# GeoGebra Math Apps
Solve math problems, graph functions, create geometric constructions, do statistics and calculus, save and share your results. All of that in your browser using [GeoGebra Graphing Calculator](http://www.geogebra.org/graphing), [GeoGebra Classic](http://www.geogebra.org/classic) and [other apps](https://www.geogebra.org/download).

# Embedding Math Apps in your website
With our JavaScript library it's easy to embed GeoGebra in any website, please see the following documentation
- https://geogebra.github.io/docs/reference/en/GeoGebra_Apps_Embedding/ -- how to embed and customize the Apps
- https://geogebra.github.io/docs/reference/en/GeoGebra_Apps_API/ -- JavaScript API to interact with the Apps
- https://geogebra.github.io/integration/ -- examples of Math Apps embedded in HTML, sources available from GitHub: https://github.com/geogebra/integration

# License
You are free to copy, distribute and transmit GeoGebra for non-commercial purposes. For details see https://www.geogebra.org/license

# Disabling popups
Open this file with vs code
geogebra-math-apps\GeoGebra\HTML5\5.0\web3d\7D43D3027ABD506A5971FC4EA283DA00.cache.js        (Filen kan have et andet mærkeligt nr. afh af udgave)

alt+z to wordwrap
search for ”beforeunload”
Så skulle der gerne komme en linje der ca. ser således ud:  fJo='beforeunload'
Delete the text in ’’
fJo=’’

Lige efter er der en med ’unload’  det ser ikke ud til at den har betydning



Popup med valg af CAS, sandsynlighed mm. er disabled via url-parameter: perspective=graphing
