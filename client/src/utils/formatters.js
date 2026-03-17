export function formatOptionLabel(value) {
  return value.replace(/([A-Z])/g, " $1").trim();
}
