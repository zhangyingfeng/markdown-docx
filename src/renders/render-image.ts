import { ImageRun } from 'docx'
import { Tokens } from 'marked'

import { MarkdownDocx } from '../MarkdownDocx'
import { ITextAttr, MarkdownImageItem } from '../types'
import { renderText } from './render-text'

export function renderImage(render: MarkdownDocx, block: Tokens.Image, attr: ITextAttr) {
  if (render.ignoreImage) {
    return false
  }

  const image = render.findImage(block)

  if (!image || !image.type) {
    return renderText(render, `[!${block.text}](${block.href})`, attr)
  }

  const { width, height, title } = parseImageTitleSize(render, block, image)

  return new ImageRun({
    type: image.type,
    data: image.data,
    transformation: { width, height },
    altText: {
      title: title || block.text,
      description: block.text,
      name: block.text
    }
  })
}


/**
 * Parse image size from token title
 * Supports format like "600x400" or "50%x50%" in title attribute
 */
export function parseImageTitleSize(render: MarkdownDocx, block: Tokens.Image, image: MarkdownImageItem) {
  const title = block.title?.trim()

  const match = title ? title.match(/^(\d+%?)x(\d+%?)$/) : null

  if (!match) {
    const { width, height } = clampImageSize(image.width, image.height, render.options.imageMaxWidth, render.options.imageMaxHeight)
    return { width, height, title: block.title }
  }

  const width = match[1].endsWith('%') ? parseInt(match[1], 10) / 100 * image.width : parseInt(match[1], 10)
  const height = match[2].endsWith('%') ? parseInt(match[2], 10) / 100 * image.height : parseInt(match[2], 10)

  return {
    width,
    height,
    // remove title
    title: ''
  }
}

/**
 * Scales width/height down (preserving aspect ratio) so neither exceeds the
 * given max, e.g. so a full-resolution screenshot doesn't render wider than
 * the page's content area. Never upscales a smaller image. An explicit
 * "WxH" title override (see parseImageTitleSize above) bypasses this
 * entirely — this only applies to an image's natural/adapter-reported size.
 */
export function clampImageSize(width: number, height: number, maxWidth?: number, maxHeight?: number) {
  const scale = Math.min(1, maxWidth ? maxWidth / width : 1, maxHeight ? maxHeight / height : 1)
  if (scale >= 1) return { width, height }
  return { width: Math.round(width * scale), height: Math.round(height * scale) }
}
