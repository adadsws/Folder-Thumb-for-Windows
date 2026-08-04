# 固定上游引用

本目录保存项目维护时参考的只读上游仓库。三个 submodule 均锁定完整 commit，不跟随浮动分支；项目运行不依赖这些仓库。

| 路径 | 上游 | 锁定 commit | Tag | 许可证 |
|---|---|---|---|---|
| `PowerToys/` | `https://github.com/microsoft/PowerToys.git` | `fc680d350f74f1f4eec8a64296420e614724e74d` | 无精确 tag | MIT |
| `mattpocock-skills/` | `https://github.com/mattpocock/skills.git` | `ed37663cc5fbef691ddfecd080dff42f7e7e350d` | 无精确 tag | MIT |
| `winutil/` | `https://github.com/ChrisTitusTech/winutil.git` | `50be7390e586f664df76d7fed41fc3c39252288c` | `26.07.22` | MIT |

从项目根目录重建：

```powershell
git submodule update --init --recursive
```

验证锁定版本：

```powershell
git submodule status --recursive
```

更新上游版本需要先确认，随后记录新的完整 SHA、tag（如有）和许可证，再提交父仓库中的 Git link 变更。
