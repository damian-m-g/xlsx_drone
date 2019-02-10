cmake_minimum_required(VERSION 3.9)
project(porcupine C)

set(CMAKE_C_STANDARD 11)

add_library(porcupine SHARED miniz.h zip.c zip.h sxmlc.c sxmlc.h sxmlsearch.c sxmlsearch.h library.c library.h)