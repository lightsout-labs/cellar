export type Language = 'en' | 'zh';

const en = {
  ribbon: 'Ribbon',
  'tab.home': 'Home',
  'tab.home.tooltip': 'The most used commands',
  undo: 'Undo',
  'undo.tooltip': 'Undo the last change',
  redo: 'Redo',
  'redo.tooltip': 'Redo the last change undone',
  'tab.view': 'View',
  'tab.view.tooltip': 'Appearance and language',
  theme: 'Theme',
  'theme.light': 'Light',
  'theme.light.tooltip': 'Use the light theme',
  'theme.dark': 'Dark',
  'theme.dark.tooltip': 'Use the dark theme',
  language: 'Language',
  'language.en.tooltip': 'Show the interface in English',
  'language.zh.tooltip': 'Show the interface in Simplified Chinese',
  nameBox: 'Name box',
  formulaBar: 'Formula bar',
  'formulaBar.tooltip': 'Contents of the active cell',
  sheets: 'Sheets',
  'sheet.tooltip': 'Double-click to rename, right-click for more actions',
  'sheet.new': 'New sheet',
  'sheet.new.tooltip': 'Insert a new sheet',
  'sheet.insert': 'Insert',
  'sheet.duplicate': 'Duplicate',
  'sheet.rename': 'Rename',
  'sheet.delete': 'Delete',
  'sheet.color': 'Tab color',
  'sheet.name': 'Sheet name',
  'sheetName.empty': 'A sheet name cannot be blank.',
  'sheetName.tooLong': 'A sheet name can have at most 31 characters.',
  'sheetName.character':
    "A sheet name cannot contain : \\ / ? * [ ] or begin or end with '.",
  'sheetName.taken': 'Another sheet already has this name.',
  'color.none': 'No color',
  'color.red': 'Red',
  'color.orange': 'Orange',
  'color.yellow': 'Yellow',
  'color.green': 'Green',
  'color.teal': 'Teal',
  'color.blue': 'Blue',
  'color.purple': 'Purple',
  'color.gray': 'Gray',
  'status.ready': 'Ready',
  'status.enter': 'Enter',
  'status.edit': 'Edit',
  'status.endMode': 'End Mode',
  summary: 'Summary of the selection',
  'summary.average': 'Average',
  'summary.count': 'Count',
  'summary.sum': 'Sum',
  zoom: 'Zoom',
  'zoom.tooltip': 'Size of the grid',
  'zoom.in': 'Zoom in',
  'zoom.in.tooltip': 'Zoom in by 10%',
  'zoom.out': 'Zoom out',
  'zoom.out.tooltip': 'Zoom out by 10%',
};

/** The name of a string of the interface. */
export type Key = keyof typeof en;

const zh: Record<Key, string> = {
  ribbon: '功能区',
  'tab.home': '开始',
  'tab.home.tooltip': '最常用的命令',
  undo: '撤消',
  'undo.tooltip': '撤消上一次更改',
  redo: '恢复',
  'redo.tooltip': '恢复上一次撤消的更改',
  'tab.view': '视图',
  'tab.view.tooltip': '外观和语言',
  theme: '主题',
  'theme.light': '浅色',
  'theme.light.tooltip': '使用浅色主题',
  'theme.dark': '深色',
  'theme.dark.tooltip': '使用深色主题',
  language: '语言',
  'language.en.tooltip': '以英文显示界面',
  'language.zh.tooltip': '以简体中文显示界面',
  nameBox: '名称框',
  formulaBar: '编辑栏',
  'formulaBar.tooltip': '活动单元格的内容',
  sheets: '工作表',
  'sheet.tooltip': '双击可重命名，右键单击可查看更多操作',
  'sheet.new': '新建工作表',
  'sheet.new.tooltip': '插入新工作表',
  'sheet.insert': '插入',
  'sheet.duplicate': '创建副本',
  'sheet.rename': '重命名',
  'sheet.delete': '删除',
  'sheet.color': '工作表标签颜色',
  'sheet.name': '工作表名称',
  'sheetName.empty': '工作表名称不能为空。',
  'sheetName.tooLong': '工作表名称最多只能有 31 个字符。',
  'sheetName.character':
    "工作表名称不能包含 : \\ / ? * [ ]，也不能以 ' 开头或结尾。",
  'sheetName.taken': '已有其他工作表使用此名称。',
  'color.none': '无颜色',
  'color.red': '红色',
  'color.orange': '橙色',
  'color.yellow': '黄色',
  'color.green': '绿色',
  'color.teal': '青色',
  'color.blue': '蓝色',
  'color.purple': '紫色',
  'color.gray': '灰色',
  'status.ready': '就绪',
  'status.enter': '输入',
  'status.edit': '编辑',
  'status.endMode': '结束模式',
  summary: '选定区域的汇总',
  'summary.average': '平均值',
  'summary.count': '计数',
  'summary.sum': '求和',
  zoom: '缩放',
  'zoom.tooltip': '网格的显示比例',
  'zoom.in': '放大',
  'zoom.in.tooltip': '放大 10%',
  'zoom.out': '缩小',
  'zoom.out.tooltip': '缩小 10%',
};

/** Every string of the interface, in each language. */
export const dictionaries: Record<Language, Record<Key, string>> = { en, zh };

/** The function that gives the strings of one language. */
export function translator(language: Language): (key: Key) => string {
  return (key) => dictionaries[language][key];
}

/**
 * The language to start in: the remembered choice, or else Chinese for a
 * browser set to any Chinese locale and English for every other.
 */
export function startingLanguage(
  remembered: string | null,
  browserLanguage: string,
): Language {
  if (remembered === 'en' || remembered === 'zh') return remembered;
  return browserLanguage.toLowerCase().startsWith('zh') ? 'zh' : 'en';
}
