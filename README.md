## kms 激活 Windows 和 office 教程

教程截至有效时间: 2019.9.21

### 相关 GitHub 链接

[kms server](https://github.com/SystemRage)
[office 转 vol 工具](https://github.com/kkkgo/office-C2R-to-VOL)

### 工具准备

0. 到[http://msdn.itellyou.cn/](http://msdn.itellyou.cn/)下载镜像并安装原版 Windows 和 office 2019
1. Python3 (我是 3.7)
1. 下载上面两个工程
1. 开干

### kms server 环境部署

PS: 不想部署自己 kms 服务器的请跳过这一段(可以使用我的 kms
服务器, 但不保证什么时候会停掉)

注意: 因为[链接](https://github.com/SystemRage/py-kms/issues/23)

```
I think the problem is "auto-activation" (py-kms / vlmcsd server running on the same PC that should be activated).
```

所以, 建议部署在有固定 ip 的服务器里, 方式很简单, 开个后台(例如`screen`), 运行`python pykms_Server.py`, 运行成功不会有信息提示, 这时部署完成, 暂且放下

### Windows10 激活

注意事项:

1. `powershell` 或者 `cmd` 管理员权限运行以下代码:
2. `W269N-WFGWX-YVC9B-4J6C9-T83GX` 为 `Windows 10 Professional`密钥, 见[链接](https://github.com/SystemRage/py-kms/wiki/Windows-GVLK-Keys#windows-10)
3. [官方教程链接](https://github.com/SystemRage/py-kms/wiki/Manual#windows)
4. 47.245.58.116 为我 kms 服务器 ip, 大家有部署自己服务器的就填对应的 IP 地址, 注意放行 1688 端口(默认, 要修改的 kms server 运行时也要指定端口)
5. `错误: 0x8007000D 数据无效。`, 该错误为在本地部署`server`程序出现的错误, 理由见[链接](https://github.com/SystemRage/py-kms/issues/23)

依次执行以下代码

```shell
cscript.exe slmgr.vbs /upk
cscript.exe //nologo slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
cscript.exe //nologo slmgr.vbs /skms 47.245.58.116:1688.
cscript.exe //nologo slmgr.vbs /ato
```

看提示, 理论上已经激活成功了
![Windows 10 activation](/images/Windows10-activation.png)

### office 2019 激活

此处教程跟官方原版教程不一致, 因为我 office 版本的问题, 需要转 VOL

`ERROR DESCRIPTION: The Software Licensing Service reported that the product SKU is not found.`错误信息为我按照官方教程配置时发现是 Retail 版本 office, 需要转为 VOL 版本, 见下步骤 1

[官方文档链接](https://github.com/SystemRage/py-kms/wiki/Manual#office)

1. 打开 office 转 vol 工具, 右键管理员权限运行`Convert-C2R.cmd`, 会提示成功换成 VOL

2. 执行以下代码

```shell
cscript.exe //nologo ospp.vbs /sethst:47.245.58.116
cscript.exe //nologo ospp.vbs /setprt:1688
cscript.exe //nologo ospp.vbs /act
```

这里应该已经提示激活成功了, 看一下激活信息

```shell
cscript.exe //nologo ospp.vbs /dstatus
```

![Office 2019 activation](/images/office2019-activation.png)

大功告成!
