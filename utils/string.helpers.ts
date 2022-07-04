function capitalize(s: string) {
  return s.toLowerCase().replace(/^\w/, (c) => c.toUpperCase());
}

export {
  capitalize
}