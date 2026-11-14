export type Language = 'en' | 'zh';

const en = {
  ribbon: 'Ribbon',
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
  sheets: 'Sheets',
  'status.ready': 'Ready',
  'status.enter': 'Enter',
};

/** The name of a string of the interface. */
export type Key = keyof typeof en;

const zh: Record<Key, string> = {
  ribbon: '功能区',
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
  sheets: '工作表',
  'status.ready': '就绪',
  'status.enter': '输入',
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
