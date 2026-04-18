/**
 * PPTX comment lifecycle commands. PowerPoint stores comments per
 * slide in `ppt/comments/commentN.xml` with a workbook-wide author
 * registry at `ppt/commentAuthors.xml`. The first add on a deck
 * synthesises the authors part, the per-slide part, and the rels +
 * content-type wiring; subsequent ops just mutate the typed list.
 *
 * IDs are derived from the OOXML primary key `${authorId}:${idx}` so
 * commands and the UI can refer to a comment without minting GUIDs that
 * wouldn't survive a serialise → re-parse cycle.
 */

import type { CommandHandler } from "@officeai/core";
import type {
  ContentTypesSnap,
  PptxComment,
  PptxCommentAuthor,
  PptxCommentAuthorsPart,
  PptxCommentsPart,
  PptxPresentation,
  PptxSnapshot,
  Slide,
} from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, makeError } from "./helpers.js";
import type {
  AddCommentPayload,
  DeleteCommentPayload,
  EditCommentPayload,
  ReplyCommentPayload,
  ResolveCommentPayload,
} from "./payloads.js";

const COMMENT_AUTHORS_PART = "ppt/commentAuthors.xml";
const COMMENT_AUTHORS_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
const COMMENTS_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
const COMMENT_AUTHORS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";

const PRESENTATION_PART = "ppt/presentation.xml";

interface CommentMutation {
  readonly snapshot: PptxSnapshot;
  readonly slide: Slide;
  readonly slideIndex: number;
}

export const addCommentHandler: CommandHandler<AddCommentPayload, PptxSnapshot> = {
  type: "pptx:add-comment",
  apply(snapshot, payload) {
    if (!payload.author || !payload.text) {
      throw makeError("invalid-payload", "author and text are required");
    }
    const ctx = locateMutation(snapshot, payload.slideIndex);
    return commitNewComment(ctx, {
      author: payload.author,
      text: payload.text,
      xEmu: payload.xEmu ?? Math.round(snapshot.root.slideSize.cxEmu / 2),
      yEmu: payload.yEmu ?? Math.round(snapshot.root.slideSize.cyEmu / 2),
    });
  },
};

export const replyCommentHandler: CommandHandler<ReplyCommentPayload, PptxSnapshot> = {
  type: "pptx:reply-comment",
  apply(snapshot, payload) {
    if (!payload.parentId || !payload.author || !payload.text) {
      throw makeError("invalid-payload", "parentId, author, text are required");
    }
    const ctx = locateMutation(snapshot, payload.slideIndex);
    const part = currentCommentsPart(ctx.slide, snapshot);
    const parent = part?.comments.find((c) => c.id === payload.parentId);
    if (!parent) throw makeError("unknown-target", `Unknown parent comment ${payload.parentId}`);
    if (parent.parentId) {
      throw makeError("invalid-payload", "Cannot reply to a reply — replies must target a top-level comment.");
    }
    // Replies adopt the parent's pin so they cluster correctly.
    return commitNewComment(
      ctx,
      {
        author: payload.author,
        text: payload.text,
        xEmu: parent.xEmu,
        yEmu: parent.yEmu,
        parentId: parent.id,
      }
    );
  },
};

export const resolveCommentHandler: CommandHandler<ResolveCommentPayload, PptxSnapshot> = {
  type: "pptx:resolve-comment",
  apply(snapshot, payload) {
    const ctx = locateMutation(snapshot, payload.slideIndex);
    const part = currentCommentsPart(ctx.slide, snapshot);
    if (!part) throw makeError("unknown-target", "Slide has no comments part");
    const found = part.comments.find((c) => c.id === payload.commentId);
    if (!found) throw makeError("unknown-target", `Unknown comment ${payload.commentId}`);
    const updated: PptxCommentsPart = {
      ...part,
      comments: part.comments.map((c) =>
        c.id === payload.commentId ? { ...c, resolved: payload.resolved } : c
      ),
    };
    return commitCommentsPartUpdate(ctx, updated, {
      kind: "node-updated",
      nodeId: ctx.slide.id,
      path: ["slides", ctx.slideIndex, "comments", payload.commentId],
      field: "resolved",
      summary: payload.resolved ? "resolved" : "reopened",
    });
  },
};

export const deleteCommentHandler: CommandHandler<DeleteCommentPayload, PptxSnapshot> = {
  type: "pptx:delete-comment",
  apply(snapshot, payload) {
    const ctx = locateMutation(snapshot, payload.slideIndex);
    const part = currentCommentsPart(ctx.slide, snapshot);
    if (!part) throw makeError("unknown-target", "Slide has no comments part");
    const found = part.comments.find((c) => c.id === payload.commentId);
    if (!found) throw makeError("unknown-target", `Unknown comment ${payload.commentId}`);
    // Cascade: if we're deleting a top-level comment, drop its replies too.
    const remaining = part.comments.filter((c) => {
      if (c.id === payload.commentId) return false;
      if (!found.parentId && c.parentId === payload.commentId) return false;
      return true;
    });
    const updated: PptxCommentsPart = { ...part, comments: remaining };
    return commitCommentsPartUpdate(ctx, updated, {
      kind: "node-deleted",
      nodeId: ctx.slide.id,
      path: ["slides", ctx.slideIndex, "comments", payload.commentId],
      summary: "comment-deleted",
    });
  },
};

export const editCommentHandler: CommandHandler<EditCommentPayload, PptxSnapshot> = {
  type: "pptx:edit-comment",
  apply(snapshot, payload) {
    if (typeof payload.text !== "string") {
      throw makeError("invalid-payload", "text must be a string");
    }
    const ctx = locateMutation(snapshot, payload.slideIndex);
    const part = currentCommentsPart(ctx.slide, snapshot);
    if (!part) throw makeError("unknown-target", "Slide has no comments part");
    const updated: PptxCommentsPart = {
      ...part,
      comments: part.comments.map((c) =>
        c.id === payload.commentId ? { ...c, text: payload.text } : c
      ),
    };
    return commitCommentsPartUpdate(ctx, updated, {
      kind: "node-updated",
      nodeId: ctx.slide.id,
      path: ["slides", ctx.slideIndex, "comments", payload.commentId],
      field: "text",
      summary: `text:${payload.text.length}ch`,
    });
  },
};

// ─── Internals ────────────────────────────────────────────────────────────

function locateMutation(snapshot: PptxSnapshot, slideIndex: number): CommentMutation {
  const { slide, index } = findSlide(snapshot, slideIndex);
  return { snapshot, slide, slideIndex: index };
}

function currentCommentsPart(slide: Slide, snapshot: PptxSnapshot): PptxCommentsPart | null {
  if (!slide.commentsPartPath) return null;
  return snapshot.root.commentsByPart.get(slide.commentsPartPath) ?? null;
}

interface NewCommentInput {
  readonly author: string;
  readonly text: string;
  readonly xEmu: number;
  readonly yEmu: number;
  readonly parentId?: string;
}

function commitNewComment(ctx: CommentMutation, input: NewCommentInput) {
  const { snapshot, slide, slideIndex } = ctx;
  const { authors, authorPart, dirtyAuthors, addedAuthorPart } = upsertAuthor(
    snapshot.root.commentAuthors,
    input.author
  );
  const author = authors.find((a) => a.name === input.author)!;
  const idx = (author.lastIdx ?? 0) + 1;
  const newAuthorList: PptxCommentAuthor[] = authors.map((a) =>
    a.id === author.id ? { ...a, lastIdx: idx } : a
  );
  const updatedAuthorsPart: PptxCommentAuthorsPart = {
    ...authorPart,
    authors: newAuthorList,
  };

  const commentsPath = slide.commentsPartPath ?? defaultCommentsPathFor(slide);
  const existingPart =
    snapshot.root.commentsByPart.get(commentsPath) ??
    ({ partPath: commentsPath, comments: [] as PptxComment[] } as PptxCommentsPart);

  const newComment: PptxComment = {
    id: `${author.id}:${idx}`,
    authorId: author.id,
    idx,
    createdAt: new Date().toISOString(),
    xEmu: input.xEmu,
    yEmu: input.yEmu,
    text: input.text,
    ...(input.parentId ? { parentId: input.parentId } : {}),
  };
  const updatedCommentsPart: PptxCommentsPart = {
    ...existingPart,
    comments: [...existingPart.comments, newComment],
  };

  // Wire rels / content-types if this is the first comment on the slide
  // and/or the first comment in the deck (authors part).
  let relationships = new Map(snapshot.relationships);
  let contentTypes = snapshot.contentTypes;
  let dirtyContentTypes = false;
  const dirtyRels: string[] = [];
  let updatedSlide = slide;

  if (!slide.commentsPartPath) {
    const slideRelsPath = relsPathFor(slide.partPath);
    const entries = [...(relationships.get(slideRelsPath)?.entries ?? [])];
    const relId = nextRelId(entries.map((e) => e.id));
    entries.push({
      id: relId,
      type: COMMENTS_REL_TYPE,
      target: relativeFrom(slideRelsPath, commentsPath),
    });
    relationships.set(slideRelsPath, { relsPath: slideRelsPath, entries });
    dirtyRels.push(slideRelsPath);
    if (!snapshot.contentTypes.overrides.some((o) => o.partName === `/${commentsPath}`)) {
      contentTypes = appendOverride(contentTypes, commentsPath, COMMENTS_CONTENT_TYPE);
      dirtyContentTypes = true;
    }
    updatedSlide = { ...slide, commentsPartPath: commentsPath };
  }

  if (addedAuthorPart) {
    // Wire presentation → commentAuthors rel.
    const presRelsPath = "ppt/_rels/presentation.xml.rels";
    const entries = [...(relationships.get(presRelsPath)?.entries ?? [])];
    if (!entries.some((e) => e.type === COMMENT_AUTHORS_REL_TYPE)) {
      const relId = nextRelId(entries.map((e) => e.id));
      entries.push({
        id: relId,
        type: COMMENT_AUTHORS_REL_TYPE,
        target: relativeFrom(presRelsPath, COMMENT_AUTHORS_PART),
      });
      relationships.set(presRelsPath, { relsPath: presRelsPath, entries });
      dirtyRels.push(presRelsPath);
    }
    if (
      !snapshot.contentTypes.overrides.some((o) => o.partName === `/${COMMENT_AUTHORS_PART}`)
    ) {
      contentTypes = appendOverride(contentTypes, COMMENT_AUTHORS_PART, COMMENT_AUTHORS_CONTENT_TYPE);
      dirtyContentTypes = true;
    }
  }

  const commentsByPart = new Map(snapshot.root.commentsByPart);
  commentsByPart.set(commentsPath, updatedCommentsPart);

  const slides = snapshot.root.slides.map((s, i) =>
    i === slideIndex ? updatedSlide : s
  );
  const root: PptxPresentation = {
    ...snapshot.root,
    slides,
    commentsByPart,
    commentAuthors: updatedAuthorsPart,
  };

  const next = evolveSnapshot(
    snapshot,
    root,
    {
      slides: updatedSlide === slide ? [] : [slide.partPath],
      comments: [commentsPath],
      commentAuthors: dirtyAuthors,
      relationships: dirtyRels,
      contentTypes: dirtyContentTypes,
    },
    { relationships, contentTypes }
  );

  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, {
      kind: "node-inserted",
      nodeId: slide.id,
      path: ["slides", slideIndex, "comments", newComment.id],
      summary: `comment:${input.text.length}ch`,
    }),
  };
}

function commitCommentsPartUpdate(
  ctx: CommentMutation,
  updatedPart: PptxCommentsPart,
  change: import("@officeai/core").DiffChange
) {
  const { snapshot, slideIndex } = ctx;
  const commentsByPart = new Map(snapshot.root.commentsByPart);
  commentsByPart.set(updatedPart.partPath, updatedPart);
  const root: PptxPresentation = { ...snapshot.root, commentsByPart };
  const next = evolveSnapshot(snapshot, root, { comments: [updatedPart.partPath] });
  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, change),
  };
}

interface UpsertResult {
  readonly authors: ReadonlyArray<PptxCommentAuthor>;
  readonly authorPart: PptxCommentAuthorsPart;
  readonly dirtyAuthors: boolean;
  readonly addedAuthorPart: boolean;
}

function upsertAuthor(
  current: PptxCommentAuthorsPart | null,
  authorName: string
): UpsertResult {
  if (!current) {
    const author: PptxCommentAuthor = { id: 0, name: authorName, lastIdx: 0 };
    return {
      authors: [author],
      authorPart: { partPath: COMMENT_AUTHORS_PART, authors: [author] },
      dirtyAuthors: true,
      addedAuthorPart: true,
    };
  }
  const existing = current.authors.find((a) => a.name === authorName);
  if (existing) {
    return {
      authors: current.authors,
      authorPart: current,
      dirtyAuthors: true,
      addedAuthorPart: false,
    };
  }
  const nextId =
    current.authors.reduce((m, a) => (a.id >= m ? a.id + 1 : m), 0) || 0;
  const newAuthor: PptxCommentAuthor = { id: nextId, name: authorName, lastIdx: 0 };
  return {
    authors: [...current.authors, newAuthor],
    authorPart: { ...current, authors: [...current.authors, newAuthor] },
    dirtyAuthors: true,
    addedAuthorPart: false,
  };
}

function defaultCommentsPathFor(slide: Slide): string {
  const m = /slide(\d+)\.xml$/i.exec(slide.partPath);
  const suffix = m ? m[1] : "1";
  return `ppt/comments/comment${suffix}.xml`;
}

function relsPathFor(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = partPath.slice(0, lastSlash);
  const file = partPath.slice(lastSlash + 1);
  return `${dir}/_rels/${file}.rels`;
}

function nextRelId(existing: ReadonlyArray<string>): string {
  let max = 0;
  for (const id of existing) {
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      const n = Number(m[1]);
      if (n > max) max = n;
    }
  }
  return `rId${max + 1}`;
}

function relativeFrom(relsPath: string, targetPath: string): string {
  const m = /^(.*?)_rels\/[^/]+\.rels$/.exec(relsPath);
  const ownerDir = (m?.[1] ?? "").replace(/\/$/, "");
  const target = targetPath.split("/");
  const owner = ownerDir.split("/").filter((s) => s.length > 0);
  let i = 0;
  while (
    i < owner.length &&
    i < target.length - 1 &&
    owner[i] === target[i]
  ) {
    i++;
  }
  const ups = owner.length - i;
  const downs = target.slice(i);
  return [...Array(ups).fill(".."), ...downs].join("/");
}

function appendOverride(
  ct: ContentTypesSnap,
  partName: string,
  contentType: string
): ContentTypesSnap {
  return {
    ...ct,
    overrides: [...ct.overrides, { partName: `/${partName}`, contentType }],
  };
}

// Suppress "unused" until presentation rels usage in serializer wires up.
void PRESENTATION_PART;
