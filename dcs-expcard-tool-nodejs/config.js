/**
 * DCS Converter - 配置文件
 * 逻辑分隔符与特殊分隔符字典（可外部修改）
 */

const DCS_CONFIG = {
  // 逻辑运算符字典
  LOGIC_OPERATORS: {
    与: "且",
    且: "且",
    或: "或",
    或取反: "或取反",
  },

  // 特殊分隔符字典
  SPECIAL_SEPARATORS: {
    或延时: "或",
    与延时: "与",
  },

  // 跳过的表头
  SKIP_HEADERS: ["序号", "条件确认"],

  // 段落标题
  SECTION_HEADERS: ["试验条件", "试验恢复", "结论", "存在问题"],
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = DCS_CONFIG;
}
