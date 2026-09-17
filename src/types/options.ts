import { IPropertiesOptions } from 'docx'
import { MarkedOptions } from 'marked'

import { MarkdownImageAdapter } from './image'
import { IMarkdownTheme } from './theme'

export interface MarkdownDocxOptions extends MarkedOptions {
  imageAdapter?: MarkdownImageAdapter

  /**
   * Math engine configuration
   * builtin: simple unicode mapping
   * katex: KaTeX -> MathML -> docx Math
   */
  math?: {
    engine?: 'builtin' | 'katex' // future: 'mathjax'
    katexOptions?: Record<string, any>
    /** Prefer constructs that are broadly supported by LibreOffice (e.g., avoid true OMML matrices and n-ary) */
    libreOfficeCompat?: boolean
  }

  /**
   * do not download image
   * @default false
   */
  ignoreImage?: boolean

  /**
   * do not parse footnote
   * @default false
   */
  ignoreFootnote?: boolean

  /**
   * do not parse html
   * @default false
   */
  ignoreHtml?: boolean

  /**
   * Strip a leading YAML frontmatter block (--- ... ---) before conversion,
   * so it doesn't get rendered as visible document text.
   * @default true
   */
  stripFrontmatter?: boolean

  /**
   * Maximum image width in pixels (at 96dpi, matching MarkdownImageItem's
   * existing width/height convention). Images wider than this are scaled
   * down proportionally; never upscaled. Ignored for an image whose
   * Markdown title already specifies an explicit size (see
   * parseImageTitleSize). Set to a falsy value to disable.
   * @default 600
   */
  imageMaxWidth?: number

  /**
   * Maximum image height in pixels, same semantics as imageMaxWidth.
   * @default 800
   */
  imageMaxHeight?: number

  /**
   * Properties for the document
   */
  document?: Omit<IPropertiesOptions, 'sections'>

  /**
   * colors theme
   */
  theme?: Partial<IMarkdownTheme>
}
