// Excel for Mac occasionally rejects otherwise-supported rich export calls
// with a bare GeneralException before anything reaches the relay. Keep this
// deliberately tiny and string-based: Office.PlatformType is not present in
// every test or older host object, while the documented platform value is.
export function isMacExcel(): boolean {
  return (Office.context?.platform as unknown) === "Mac";
}
