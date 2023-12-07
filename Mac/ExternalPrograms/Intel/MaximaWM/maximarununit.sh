#!/usr/bin/env zsh
# Denne fil er til for at kunne starte Maxima helt normalt bla for at kunne lave maxima.core og andet test. Den skal ligge i MaximaWM
# den startes med 'sh maxima.sh'
#SCRIPT="${BASH_SOURCE[0]}"
#while [ -L "$SCRIPT" ] ; do SCRIPT=`(readlink "$SCRIPT")` ; done

#ROOT=`(cd \`dirname "$SCRIPT"\` > /dev/null 2>&1 ; pwd)`
#MAXIMA_PREFIX=$ROOT/maxima/
MAXIMA_PREFIX='/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/WordMat/MaximaWM/maxima'
export MAXIMA_PREFIX

PATH="$MAXIMA_PREFIX/bin:$PATH"
export PATH
#function timeout() { perl -e 'alarm shift; exec @ARGV' "$@"; }
echo Starting Maxima M1-compiled

#exec "$MAXIMA_PREFIX/bin/maxima" "$1" "$2" "$3" "$3" "$4" "$5" "$6" "$7" "$8" "$9"
# "$MAXIMA_PREFIX/bin/Maximatimeout" 10
"$MAXIMA_PREFIX/bin/sbcl" --core "$MAXIMA_PREFIX/lib/maxima/5.47.0/binary-sbcl/maximaunit.core"
