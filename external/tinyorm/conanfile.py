from conan import ConanFile
from conan.tools.cmake import CMake, CMakeToolchain, CMakeDeps, cmake_layout
from conan.tools.scm import Git
from conan.tools.files import copy

import os

class TinyOrmConan(ConanFile):
    name = "tinyorm"
    version = "0.38.1"
    url = "https://github.com/silverqx/TinyORM"
    license = "MIT"
    description = "Modern C++ ORM library"
    settings = "os", "compiler", "build_type", "arch"
    tool_requires = "cmake/[>=3.27]"

    options = {"shared": [True, False]}
    default_options = {"shared": False}

    def layout(self):
        cmake_layout(self)

    def requirements(self):
        self.requires("range-v3/0.12.0", transitive_headers=True, transitive_libs=True)
        self.requires("tabulate/1.5", transitive_headers=True, transitive_libs=True)

    def generate(self):
        qt_path = os.environ.get("QT_INSTALL_DIR", "C:/Qt6/6.10.0/msvc2022_64")
        tc = CMakeToolchain(self)
        tc.variables["BUILD_SHARED_LIBS"] = "ON" if self.options.shared else "OFF"
        tc.variables["TOM"] = "ON"
        tc.variables["ORM"] = "ON"
        tc.generate()

        deps = CMakeDeps(self)
        deps.generate()

    def source(self):
        git = Git(self)
        git.clone("https://github.com/silverqx/TinyORM", target=".", args=["--branch", "v0.38.1", "--depth", "1"])

    def build(self):
        cmake = CMake(self)
        qt_path = "C:/Qt6/6.10.0/msvc2022_64"
        cmake.configure(variables={
            "CMAKE_PREFIX_PATH": qt_path
        })
        cmake.build()

    def package(self):
        cmake = CMake(self)
        cmake.install()
        copy(self, "LICENSE", src=self.source_folder, dst=self.package_folder + "/licenses")

    def package_info(self):
        self.cpp_info.set_property("cmake_file_name", "tinyorm")
        self.cpp_info.set_property("cmake_target_name", "tinyorm::tinyorm")
        self.cpp_info.set_property("cmake_multi", True)

        if self.settings.build_type == "Debug":
            self.cpp_info.libs = ["TinyOrm0d"]   # Debug library name
        else:
            self.cpp_info.libs = ["TinyOrm0"]    # Release / RelWithDebInfo

        self.cpp_info.includedirs = ["include"]
        self.cpp_info.defines = ["TINYORM_USING_QTSQLDRIVERS"]

        self.cpp_info.requires = ["range-v3::range-v3", "tabulate::tabulate"]

        self.cpp_info.set_property("cmake_target_transitive_headers", True)