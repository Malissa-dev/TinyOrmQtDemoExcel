cmake_minimum_required(VERSION 3.22)

# ---- Project and Global Settings ----
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

project(TinyOrmQtDemo LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_compile_definitions(
    WIN32_LEAN_AND_MEAN
    NOMINMAX
    _WIN32_WINNT=0x0A00
    WINVER=0x0A00
)

set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreadedDLL" CACHE STRING "" FORCE)

# /bigobj for debug builds to avoid "fatal error C1128: number of sections exceeded object file format limit"
if (MSVC)
    add_compile_options($<$<CONFIG:Debug>:/bigobj>)
endif()

# ---- Qt Setup ----
set(QT_INSTALL_DIR "C:/Qt6/6.10.0/msvc2022_64" CACHE PATH "Path to Qt installation")
list(APPEND CMAKE_PREFIX_PATH ${QT_INSTALL_DIR})

find_package(Qt6 REQUIRED COMPONENTS Core Widgets Sql)
qt_standard_project_setup()

# ---- Conan Dependencies ----
find_package(tinyorm REQUIRED)

# ---- OpenXSLX
add_subdirectory(external/OpenXLSX)

# ---- Sources ----
file(GLOB_RECURSE SOURCES "src/*.cpp")
qt_add_resources(RESOURCES_QRC resources/resources.qrc)

qt_add_executable(TinyOrmQtDemo WIN32
    ${SOURCES}
    ${RESOURCES_QRC}
)

target_link_libraries(TinyOrmQtDemo
    PRIVATE
        Qt6::Core
        Qt6::Widgets
        Qt6::Sql
        tinyorm::tinyorm
        OpenXLSX::OpenXLSX
)

# ---- Qt Deployment ----
set(QT_DEPLOY_BIN_DIR "${CMAKE_RUNTIME_OUTPUT_DIRECTORY}")
include("${CMAKE_BINARY_DIR}/.qt/QtDeploySupport-Release.cmake" OPTIONAL)

# deploy Qt DLLs next to the exe
add_custom_target(DeployQt ALL
    COMMAND "${__QT_DEPLOY_TOOL}"
            --dir "$<TARGET_FILE_DIR:TinyOrmQtDemo>"
            "$<TARGET_FILE:TinyOrmQtDemo>"
    COMMENT "Deploying Qt DLLs using windeployqt"
    VERBATIM
)
