import { DocumentSessionPage } from "./document-session-page";

export const dynamic = "force-dynamic";

export default async function Page({
  params,
}: {
  readonly params: Promise<{ readonly documentId: string }>;
}) {
  const { documentId } = await params;
  return <DocumentSessionPage documentId={documentId} />;
}
