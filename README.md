# TinyOrmQtDemo

**TinyOrmQtDemo** is a demonstration proof-of-concept application built with **Qt** and **TinyORM**.  
It showcases a basic integration of a Qt GUI with TinyORM models using modern C++ and CMake build tools.

---

## Features

- Simple Qt-based GUI  
- Demonstrates CRUD operations using TinyORM  
- Modern CMake and Conan2 integration  

---

## Prerequisites

- C++ compiler (MSVC 2022 recommended)  
- Windows 11 SDK  
- Qt 6.10 (MSVC 2022 64-bit) installed via Qt Installer  
- [CMake](https://cmake.org/)  
- [Conan2](https://conan.io/)  

---

## Building the Project

> These instructions assume a Windows + MSVC 2022 environment. Other compilers or OSes may require additional configuration.

1. **Install Qt 6.10** using the official installer and note the install path (e.g., `${CMAKE_SOURCE_DIR}/external/Qt/msvc2022_64`).  


2. **Install Conan2** if not already installed:

```
pip install conan
```

3. **Set up a Conan profile** for your compiler and environment:

```
conan profile detect
```

4. **Prepare TinyORM dependencies**:  
   Navigate to the `external/tinyorm` directory and run:

```
conan create . -s build_type=Release
conan create . -s build_type=Debug
```

> These commands build TinyORM in both Release and Debug configurations so your application can link correctly. This only has to be done once per TinyORM version/build.


5. **Install dependencies in the root project**:  
   From the root of the repository, run:

```
conan install . --build=missing -s build_type=Release
conan install . --build=missing -s build_type=RelWithDebInfo
conan install . --build=missing -s build_type=Debug
```

6. **Configure the project with CMake**:

```
cmake --preset app-config-base
```

7. **Build the project**:

```
cmake --build --preset Release
cmake --build --preset RelWithDebInfo
cmake --build --preset Debug
```

8. **Deploy Qt DLLs and plugins** next to the executable:

```
cmake --build --preset Release --target DeployQt
cmake --build --preset RelWithDebInfo --target DeployQt
cmake --build --preset Debug --target DeployQt
```

9. **Executable location**:  
   After building and deploying, the `.exe` will be in the `build` folder, along with the required Qt DLLs.

---

## Example Conan Profile

This project was built using the following Conan profile:

```
[settings]
arch=x86_64
build_type=Release
compiler=msvc
compiler.cppstd=20
compiler.runtime=dynamic
compiler.version=194
os=Windows

[conf]
tools.cmake.cmaketoolchain:generator=Ninja Multi-Config
```

> Compatibility with other profiles is not guaranteed; adjustments may be required.

---

## License

This application is licensed under the **MIT License**.

### Third-party Dependencies

- [Qt](https://www.qt.io/) — LGPL 3.0  
- [TinyORM](https://github.com/fnc12/tinyorm) — MIT License
