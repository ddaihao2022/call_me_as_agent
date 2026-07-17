const segmenter = typeof Intl !== 'undefined' && 'Segmenter' in Intl
  ? new Intl.Segmenter(undefined, { granularity: 'grapheme' })
  : null

// Split text for simulated streaming without breaking surrogate pairs
// (e.g. emoji) across chunks. Prefers grapheme clusters so ZWJ sequences
// and flag emoji also stay intact; falls back to code points.
export const splitGraphemes = (text: string): string[] => {
  if (segmenter) {
    return Array.from(segmenter.segment(text), s => s.segment)
  }
  return Array.from(text)
}
