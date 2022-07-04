function capitalize(s: string) {
  return s.toLowerCase().replace(/^\w/, (c) => c.toUpperCase());
}

function escapeControlCodes(s: string) {
  return JSON.stringify({s}).match(/^{"s":\s*"(?<s>.*)"}$/).groups.s;
}

function unescapeControlCodes(s: string) {
  return JSON.parse(`{"s": "${s}"}`).s;
}

export {
  capitalize,
  escapeControlCodes,
  unescapeControlCodes
}