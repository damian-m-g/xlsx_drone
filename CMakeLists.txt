cmake_minimum_required(VERSION 3.9)
project(xlsx_drone C)

set(CMAKE_C_STANDARD 11)

add_library(xlsx_drone SHARED ext/miniz.h ext/zip.c ext/zip.h ext/sxmlc.c ext/sxmlc.h ext/sxmlsearch.c ext/sxmlsearch.h src/xlsx_drone.c src/xlsx_drone.h)