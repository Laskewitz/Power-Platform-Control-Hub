/**
 * Maps common Power Platform LCID values to BCP-47 locale tags.
 * Source: https://learn.microsoft.com/en-us/power-platform/admin/language-collations
 */
const LCID_TO_BCP47: Record<number, string> = {
  1025: 'ar-SA',  1026: 'bg-BG',  1027: 'ca-ES',  1028: 'zh-TW',
  1029: 'cs-CZ',  1030: 'da-DK',  1031: 'de-DE',  1032: 'el-GR',
  1033: 'en-US',  1034: 'es-ES',  1035: 'fi-FI',  1036: 'fr-FR',
  1037: 'he-IL',  1038: 'hu-HU',  1040: 'it-IT',  1041: 'ja-JP',
  1042: 'ko-KR',  1043: 'nl-NL',  1044: 'nb-NO',  1045: 'pl-PL',
  1046: 'pt-BR',  1048: 'ro-RO',  1049: 'ru-RU',  1050: 'hr-HR',
  1051: 'sk-SK',  1053: 'sv-SE',  1054: 'th-TH',  1055: 'tr-TR',
  1057: 'id-ID',  1058: 'uk-UA',  1060: 'sl-SI',  1061: 'et-EE',
  1062: 'lv-LV',  1063: 'lt-LT',  1069: 'eu-ES',  1086: 'ms-MY',
  1110: 'gl-ES',  2052: 'zh-CN',  2057: 'en-GB',  2058: 'es-MX',
  2070: 'pt-PT',  3081: 'en-AU',  3082: 'es-ES',  4105: 'en-CA',
};

let displayNames: Intl.DisplayNames | null = null;
function getDisplayNames(): Intl.DisplayNames {
  if (!displayNames) {
    try {
      displayNames = new Intl.DisplayNames(undefined, { type: 'language' });
    } catch {
      displayNames = new Intl.DisplayNames(['en'], { type: 'language' });
    }
  }
  return displayNames;
}

/**
 * Converts an LCID number to a human-readable string: "English (1033)".
 * Falls back to just the number if the LCID is not in the map.
 */
export function lcidToLabel(lcid: number): string {
  const tag = LCID_TO_BCP47[lcid];
  if (!tag) return String(lcid);
  try {
    const name = getDisplayNames().of(tag);
    return name ? `${name} (${lcid})` : String(lcid);
  } catch {
    return String(lcid);
  }
}
