# VMware 虚拟机批量转 OVF 工具

这是一个 PowerShell 脚本，用于自动扫描当前目录下的所有 VMware 虚拟机配置文件 (`.vmx`)，并使用 `ovftool` 将它们批量转换为开放虚拟化格式 (`.ovf`)，以便导入到 VirtualBox 或其他虚拟化平台。

## 功能特点

- **自动发现**：递归扫描当前目录及子目录，查找所有 `.vmx` 虚拟机文件。
- **批量转换**：自动调用 VMware OVF Tool 进行格式转换。
- **目录隔离**：转换后的文件按虚拟机名称分类保存在 `_OVF_Exports` 目录下，保持整洁。
- **智能过滤**：自动排除已经导出的目录，避免重复处理。
- **错误处理**：如果某个虚拟机转换失败，脚本会记录错误并继续处理下一个虚拟机。

## 前置要求

1.  **操作系统**：Windows (PowerShell 5.1 或更高版本)。
2.  **依赖工具**：必须安装 **VMware OVF Tool**。
    - 脚本默认查找路径：`D:\Program Files\ovftool\ovftool.exe`
    - 如果您的安装路径不同，请用文本编辑器打开 `convert_vms.ps1`，修改第 2 行的 `$ovfToolPath` 变量。

## 使用方法

1.  将 `convert_vms.ps1` 脚本放置在包含虚拟机的根目录下（例如 `D:\VM`）。
2.  打开 PowerShell 终端。
3.  切换到脚本所在目录：
    ```powershell
    cd D:\VM
    ```
4.  运行脚本：
    ```powershell
    .\convert_vms.ps1
    ```
    *如果提示权限错误，请使用以下命令绕过执行策略运行：*
    ```powershell
    PowerShell -ExecutionPolicy Bypass -File .\convert_vms.ps1
    ```

## 输出结构

脚本运行后，会在当前目录下创建一个名为 `_OVF_Exports` 的文件夹。目录结构如下：

```text
D:\VM\
├── convert_vms.ps1          # 本脚本
├── CentOS 7/                # 原始虚拟机目录
│   └── CentOS 7.vmx
├── ...
└── _OVF_Exports/            # [自动创建] 导出目录
    ├── CentOS 7/
    │   ├── CentOS 7.ovf     # 转换后的 OVF 描述文件
    │   ├── CentOS 7.mf      # 校验文件
    │   └── CentOS 7-disk1.vmdk # 磁盘文件
    └── Windows 10/
        └── ...
```

## 常见问题与注意事项

1.  **转换时间**：
    转换过程涉及大量磁盘读写和格式转换，速度取决于磁盘性能（HDD/SSD）和虚拟机大小。请耐心等待，不要关闭终端窗口。

2.  **VirtualBox 兼容性**：
    生成的 `.ovf` 文件通常可以直接在 VirtualBox 中通过 **"管理" -> "导入虚拟电脑"** 打开。

3.  **macOS 虚拟机**：
    脚本会尝试转换 macOS 虚拟机，但由于 Apple 硬件许可限制，转换后的 macOS 虚拟机可能无法在非 Apple 硬件上的 VirtualBox 中启动。

4.  **路径中的空格**：
    脚本已针对包含空格的文件名（如 "CentOS 7 64 位"）进行了优化，可以正常处理。

5.  **运行报错**：
    - 如果出现红色报错文字，请检查该虚拟机是否已损坏或被占用（请确保虚拟机处于**关机**状态）。
    - 确保 `ovftool` 路径配置正确。
