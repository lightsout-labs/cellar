import { describe, expect, it } from 'vitest';
import { dictionaries, startingLanguage, translator } from './i18n';

describe('startingLanguage', () => {
  it('follows the browser: Chinese for any Chinese locale, else English', () => {
    expect(startingLanguage(null, 'zh-CN')).toBe('zh');
    expect(startingLanguage(null, 'zh')).toBe('zh');
    expect(startingLanguage(null, 'en-GB')).toBe('en');
    expect(startingLanguage(null, 'fr-FR')).toBe('en');
  });

  it('prefers the remembered choice, and ignores one it does not know', () => {
    expect(startingLanguage('en', 'zh-CN')).toBe('en');
    expect(startingLanguage('zh', 'en-US')).toBe('zh');
    expect(startingLanguage('fr', 'zh-CN')).toBe('zh');
  });
});

describe('translator', () => {
  it('gives the string of its language', () => {
    expect(translator('en')('status.ready')).toBe('Ready');
    expect(translator('zh')('status.ready')).toBe('就绪');
  });

  it('has a translation of every string, different in each language', () => {
    const { en, zh } = dictionaries;
    for (const key of Object.keys(en) as (keyof typeof en)[]) {
      expect(zh[key], key).toMatch(/\p{Script=Han}/u);
      expect(en[key], key).toMatch(/^[\x20-\x7e]+$/);
    }
  });
});
