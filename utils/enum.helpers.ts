type EnumObject = {[key: string]: number | string};

function getEnumValues<E extends EnumObject>(enumObject: E): string[] {
  return Object.keys(enumObject).filter(k => typeof enumObject[k as any] === "number");
}

export {
  getEnumValues,
}