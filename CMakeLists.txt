cmake_minimum_required(VERSION 3.9)
project(porcupine C)

set(CMAKE_C_STANDARD 11)

add_library(porcupine SHARED ext/miniz.h ext/zip.c ext/zip.h ext/sxmlc.c ext/sxmlc.h ext/sxmlsearch.c ext/sxmlsearch.h src/library.c src/library.h)