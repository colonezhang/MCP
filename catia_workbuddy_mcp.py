# -*- coding: utf-8 -*-
"""
CATIA V5 最简 MCP 服务器
功能：
  1. catia_create_sketch  - 在 XY 平面上创建草图（画一个圆）
  2. catia_pad            - 将草图拉伸成实体

依赖：
  pip install mcp pywin32

使用前提：CATIA V5 已启动并打开/新建了一个 Part 文档。
运行：python catia_mcp.py
"""
import math
import win32com.client
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("catia-v5")

_catia = None


def get_catia():
    """连接到已运行的 CATIA 实例"""
    global _catia
    if _catia is None:
        _catia = win32com.client.Dispatch("CATIA.Application")
        _catia.Visible = True
    return _catia


def get_active_part():
    """获取当前激活的 Part 文档"""
    catia = get_catia()
    doc = catia.ActiveDocument
    if doc is None or not doc.Name.lower().endswith(".catpart"):
        raise RuntimeError("请先在 CATIA 中新建或打开一个 Part 文档 (.CATPart)")
    return doc


def get_xy_plane(part):
    """获取 XY 平面"""
    return part.OriginElements.FindObjectByName("xy plane")


@mcp.tool()
def catia_create_sketch(sketch_name: str = "Sketch.1",
                        circle_x: float = 0.0,
                        circle_y: float = 0.0,
                        radius: float = 20.0) -> str:
    """在 XY 平面上创建草图并画一个圆。

    Args:
        sketch_name: 草图名称，如 Sketch.1
        circle_x: 圆心 X 坐标 (mm)
        circle_y: 圆心 Y 坐标 (mm)
        radius: 圆半径 (mm)
    """
    part = get_active_part()
    body = part.Bodies.Item(1)
    sketches = body.Sketches
    xy_plane = get_xy_plane(part)

    # 在 XY 平面上新建草图
    sketch = sketches.Add(xy_plane)
    sketch.Name = sketch_name

    # 画圆
    factory = sketch.OpenEdition()
    factory.AddClosedCircle(circle_x, circle_y, radius)
    sketch.CloseEdition()

    part.Update()
    return (f"草图 '{sketch_name}' 创建成功："
            f"圆心({circle_x}, {circle_y}) 半径 {radius} mm")


@mcp.tool()
def catia_pad(sketch_name: str = "Sketch.1",
              length: float = 20.0,
              pad_name: str = "Pad.1") -> str:
    """将指定草图沿法向拉伸成实体。

    Args:
        sketch_name: 要拉伸的草图名称
        length: 拉伸长度 (mm)
        pad_name: 拉伸特征名称
    """
    part = get_active_part()
    body = part.Bodies.Item(1)

    # 查找草图
    sketch = None
    for i in range(1, body.Sketches.Count + 1):
        s = body.Sketches.Item(i)
        if s.Name == sketch_name:
            sketch = s
            break
    if sketch is None:
        return f"错误：找不到草图 '{sketch_name}'，请先用 catia_create_sketch 创建"

    # 创建 Pad
    shape_factory = part.ShapeFactory
    pad = shape_factory.AddNewPad(sketch, length)
    pad.Name = pad_name

    part.Update()
    return f"拉伸 '{pad_name}' 创建成功：基于草图 '{sketch_name}'，长度 {length} mm"


@mcp.tool()
def catia_list_objects() -> str:
    """列出当前 Part 中的所有草图和实体特征。"""
    part = get_active_part()
    body = part.Bodies.Item(1)
    lines = []
    for i in range(1, body.Sketches.Count + 1):
        s = body.Sketches.Item(i)
        lines.append(f"[草图] {s.Name}")
    for i in range(1, body.Shapes.Count + 1):
        s = body.Shapes.Item(i)
        lines.append(f"[特征] {s.Name}")
    if not lines:
        return "当前 Part 中没有草图或特征"
    return "\n".join(lines)


if __name__ == "__main__":
    mcp.run()  # 默认 stdio 传输
