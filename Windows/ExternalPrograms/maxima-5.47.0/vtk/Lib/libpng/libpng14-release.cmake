#----------------------------------------------------------------
# Generated CMake target import file for configuration "Release".
#----------------------------------------------------------------

# Commands may need to know the format version.
set(CMAKE_IMPORT_FILE_VERSION 1)

# Import target "png14" for configuration "Release"
set_property(TARGET png14 APPEND PROPERTY IMPORTED_CONFIGURATIONS RELEASE)
set_target_properties(png14 PROPERTIES
  IMPORTED_IMPLIB_RELEASE "${_IMPORT_PREFIX}/lib/libpng14.lib"
  IMPORTED_LINK_INTERFACE_LIBRARIES_RELEASE "C:/CmbDash/pvsb/build/install/lib/zlib.lib"
  IMPORTED_LOCATION_RELEASE "${_IMPORT_PREFIX}/bin/libpng14.dll"
  )

list(APPEND _IMPORT_CHECK_TARGETS png14 )
list(APPEND _IMPORT_CHECK_FILES_FOR_png14 "${_IMPORT_PREFIX}/lib/libpng14.lib" "${_IMPORT_PREFIX}/bin/libpng14.dll" )

# Import target "png14_static" for configuration "Release"
set_property(TARGET png14_static APPEND PROPERTY IMPORTED_CONFIGURATIONS RELEASE)
set_target_properties(png14_static PROPERTIES
  IMPORTED_LINK_INTERFACE_LANGUAGES_RELEASE "C"
  IMPORTED_LINK_INTERFACE_LIBRARIES_RELEASE "C:/CmbDash/pvsb/build/install/lib/zlib.lib"
  IMPORTED_LOCATION_RELEASE "${_IMPORT_PREFIX}/lib/libpng14_static.lib"
  )

list(APPEND _IMPORT_CHECK_TARGETS png14_static )
list(APPEND _IMPORT_CHECK_FILES_FOR_png14_static "${_IMPORT_PREFIX}/lib/libpng14_static.lib" )

# Commands beyond this point should not need to know the version.
set(CMAKE_IMPORT_FILE_VERSION)
