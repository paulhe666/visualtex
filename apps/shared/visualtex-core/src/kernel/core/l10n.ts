import { isBrowser } from '../ui/utils/capabilities';

// The locale (language + region) of the editor. It selects the default letter
// shape style and the keyboard layout; the kernel has no localized UI strings.
let gLocale = '';

export const l10n = {
  get locale(): string {
    // Use the browser defined language as the default language,
    // "english" if not running in a browser (node.js)
    if (!gLocale)
      gLocale = isBrowser() ? navigator.language.slice(0, 5) : 'en-US';
    return gLocale;
  },

  set locale(value: string) {
    gLocale = value;
  },
};
