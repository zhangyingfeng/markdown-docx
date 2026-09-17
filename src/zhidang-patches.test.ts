// Tests for 知档 (zhidang, github.com/zhangyingfeng/zhi-dang)'s own additions
// to this fork — kept in a separate file from upstream's own tests so a
// future `git merge upstream/main` never has to reconcile changes to this
// file with changes to ours in the same place.
import { describe, expect, it, vi } from 'vitest'

import { MarkdownDocx } from './entry-node'
import { clampImageSize } from './renders/render-image'
import { stripFrontmatter } from './utils'

describe('stripFrontmatter', () => {
  it('removes a leading YAML frontmatter block', () => {
    const md = `---\ntitle: "hello"\nurl: https://example.com\n---\n\n# Heading\n\nBody text.`
    expect(stripFrontmatter(md)).toBe('# Heading\n\nBody text.')
  })

  it('leaves markdown with no frontmatter untouched', () => {
    const md = '# Heading\n\nBody text.'
    expect(stripFrontmatter(md)).toBe(md)
  })

  it('does not touch a "---" that is not at the very start of the document (e.g. a horizontal rule)', () => {
    const md = '# Heading\n\n---\n\nBody text.'
    expect(stripFrontmatter(md)).toBe(md)
  })
})

describe('clampImageSize', () => {
  it('scales down proportionally when width exceeds the max', () => {
    expect(clampImageSize(1600, 900, 600, 800)).toEqual({ width: 600, height: 338 })
  })

  it('scales down proportionally when height exceeds the max', () => {
    expect(clampImageSize(800, 1200, 600, 800)).toEqual({ width: 533, height: 800 })
  })

  it('never upscales an image smaller than the max', () => {
    expect(clampImageSize(100, 50, 600, 800)).toEqual({ width: 100, height: 50 })
  })

  it('is a no-op when no max is given', () => {
    expect(clampImageSize(3000, 3000, undefined, undefined)).toEqual({ width: 3000, height: 3000 })
  })
})

describe('MarkdownDocx integration', () => {
  vi.mock('docx')

  it('clamps an oversized image to the default max (600x800) when rendered', async () => {
    const docx = new MarkdownDocx('![](big.png)', {
      imageAdapter: async () => ({ type: 'png', data: Buffer.from(''), width: 1600, height: 900 }),
    })
    const rows = await docx.toSection()
    const text = rows.map(r => r.toString()).join('\n')
    expect(text).toContain('transformation-width="600"')
    expect(text).toContain('transformation-height="338"')
  })

  it('still honors an explicit "WxH" title override, bypassing the clamp', async () => {
    const docx = new MarkdownDocx('![](big.png "1200x700")', {
      imageAdapter: async () => ({ type: 'png', data: Buffer.from(''), width: 1600, height: 900 }),
    })
    const rows = await docx.toSection()
    const text = rows.map(r => r.toString()).join('\n')
    expect(text).toContain('transformation-width="1200"')
    expect(text).toContain('transformation-height="700"')
  })

  it('strips frontmatter by default so it never appears as rendered text', async () => {
    const docx = new MarkdownDocx('---\ntitle: "secret metadata"\n---\n\n# Real Heading\n')
    const rows = await docx.toSection()
    const text = rows.map(r => r.toString()).join('\n')
    expect(text).not.toContain('secret metadata')
    expect(text).toContain('Real Heading')
  })

  it('keeps frontmatter visible when stripFrontmatter is explicitly disabled', async () => {
    const docx = new MarkdownDocx('---\ntitle: "visible metadata"\n---\n\n# Real Heading\n', { stripFrontmatter: false })
    const rows = await docx.toSection()
    const text = rows.map(r => r.toString()).join('\n')
    expect(text).toContain('visible metadata')
  })
})
