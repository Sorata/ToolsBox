# VMware 虚拟机批量转 OVF 及导入 VirtualBox 工具

这是一套 PowerShell 脚本，用于实现 VMware 虚拟机到 VirtualBox 的批量迁移。它包含两个脚本：
1. `convert_vms.ps1`: 将 `.vmx` 虚拟机转换为 `.ovf` 格式。
2. `import_vms.ps1`: 将转换好的 `.ovf` 文件批量导入到 VirtualBox。

## 功能特点

- **自动发现**：递归扫描目录查找虚拟机文件。
- **批量处理**：自动调用 VMware OVF Tool 和 VirtualBox VBoxManage 进行批量转换和导入。
- **目录隔离**：转换后的文件按虚拟机名称分类保存在 `_OVF_Exports` 目录下。
- **智能过滤**：自动排除已经处理的目录。
- **日志记录**：导入过程会自动记录到 `import_log.txt`，方便排查错误。
- **自动命名**：导入时强制使用 OVF 文件名作为虚拟机名称，避免命名冲突。

## 前置要求

1.  **操作系统**：Windows (PowerShell 5.1 或更高版本)。
2.  **VMware OVF Tool**：用于转换 VM。
    - 默认路径：`D:\Program Files\ovftool\ovftool.exe`
3.  **VirtualBox**：用于导入 VM。
    - 默认路径：`D:\Program Files\Oracle\VirtualBox\VBoxManage.exe`

*如果您的安装路径不同，请分别打开脚本文件修改开头的路径变量。*

## 使用方法

### 第一步：批量转换 (VMware -> OVF)

1.  将脚本放置在包含 VMware 虚拟机的根目录下（例如 `D:\VM`）。
2.  打开 PowerShell 并运行：
    ```powershell
    .\convert_vms.ps1
    ```
    *（结果将保存在 `_OVF_Exports` 文件夹中）*

### 第二步：批量导入 (OVF -> VirtualBox)

1.  确保 `_OVF_Exports` 目录中已有转换好的 OVF 文件。
2.  在同级目录下运行导入脚本：
    ```powershell
    .\import_vms.ps1
    ```
3.  **查看日志**：导入进度和结果会实时输出到终端，并保存在 `import_log.txt` 文件中。

*如果提示权限错误，请使用以下命令绕过执行策略运行：*
```powershell
PowerShell -ExecutionPolicy Bypass -File .\convert_vms.ps1
PowerShell -ExecutionPolicy Bypass -File .\import_vms.ps1
```

## 目录结构示例

```text
D:\VM\
├── convert_vms.ps1          # 转换脚本
├── import_vms.ps1           # 导入脚本
├── import_log.txt           # [自动生成] 导入日志
├── CentOS 7/                # 原始虚拟机目录
├── ...
└── _OVF_Exports/            # [自动创建] 导出目录
    ├── CentOS 7/
    │   └── CentOS 7.ovf     # 中间文件
    └── Windows 10/
        └── ...
```

## 常见问题与故障排除

1.  **导入失败：NVMe Controller Error**
    - **现象**：报错 `Error reading ... (subtype:vmware.nvme.controller)`。
    - **原因**：VMware 使用了 VirtualBox 不支持的 NVMe 控制器定义。
    - **解决方法**：
        1. 用文本编辑器打开对应的 `.ovf` 文件。
        2. 删除包含 `vmware.nvme.controller` 的 `<Item>...</Item>` 块。
        3. 找到磁盘的 `<Item>` 定义，将其 `<rasd:Parent>` 值修改为 SATA 控制器的 InstanceID（通常是 SATA Controller 的 InstanceID，例如 `3` 或 `4`）。
        4. 删除同目录下的 `.mf` (Manifest) 校验文件，否则会报校验错误。
        5. 重新运行导入脚本。

2.  **虚拟机已存在**
    - 脚本会跳过 VirtualBox 中已存在的同名虚拟机。如需覆盖，请先手动在 VirtualBox 中删除旧虚拟机。

3.  **macOS 虚拟机**
    - 由于 Apple 硬件限制，转换后的 macOS 虚拟机可能无法在非 Apple 硬件上的 VirtualBox 中启动。

4.  **Windows 10 启动蓝屏 (INACCESSIBLE_BOOT_DEVICE)**
    - 如果在修复 NVMe 问题后将磁盘改为了 SATA，Windows 可能会因为缺少驱动而蓝屏。通常重启几次后 Windows 会自动修复驱动配置。
