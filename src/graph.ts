import { getToken } from "./auth";

export type SharePointHit = {
  title: string;
  url: string;
  summary: string;
  source: string;
};

type GraphSearchResponse = {
  value?: Array<{
    hitsContainers?: Array<{
      hits?: Array<{
        rank?: number;
        summary?: string;
        resource?: {
          name?: string;
          webUrl?: string;
          parentReference?: { siteId?: string };
        };
      }>;
    }>;
  }>;
};

function siteFilter(): string {
  const site = (import.meta.env.VITE_SHAREPOINT_SITE_URL as string | undefined)?.trim();
  if (!site) return "";
  return ` AND Path:"${site.replace(/"/g, "")}"`;
}

export async function searchSharePoint(name: string): Promise<SharePointHit[]> {
  const query = name.trim();
  if (!query) return [];

  const token = await getToken();
  const body = {
    requests: [
      {
        entityTypes: ["driveItem", "listItem"],
        query: {
          queryString: `${query}${siteFilter()}`,
        },
        from: 0,
        size: 15,
      },
    ],
  };

  const response = await fetch("https://graph.microsoft.com/v1.0/search/query", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`SharePoint search failed (${response.status}): ${detail}`);
  }

  const data = (await response.json()) as GraphSearchResponse;
  const hits: SharePointHit[] = [];

  for (const container of data.value ?? []) {
    for (const hitGroup of container.hitsContainers ?? []) {
      for (const hit of hitGroup.hits ?? []) {
        const title = hit.resource?.name ?? "Untitled";
        const url = hit.resource?.webUrl ?? "";
        if (!url) continue;
        hits.push({
          title,
          url,
          summary: (hit.summary ?? "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim(),
          source: "SharePoint",
        });
      }
    }
  }

  return hits;
}
