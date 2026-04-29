/**
 * Local type alias for the per-site inventory shape we read from the
 * results blob. Defined separately so the replace function doesn't pull
 * in the entire scanner module just for the type.
 */
export interface ISiteInventoryShape {
  site: string;
  pageCount: number;
  linkCount: number;
  pages: Array<{
    pageId: number;
    pageTitle: string;
    pageUrl: string;
    etag?: string;
    links: Array<{
      rawUrl: string;
      url: string;
      source: string;
      linkClass: string;
      normalizedKey: string;
      canonicalKey?: string;
      suggestion?: string;
      text?: string;
    }>;
  }>;
  error?: string;
}
