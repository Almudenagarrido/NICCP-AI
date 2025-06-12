import path from "path";
import fs from "fs/promises";

export default async function Page({ params }) {
  const filePath = path.join(process.cwd(), "app", "case-studies", `${params.id}`);
  const content = await fs.readFile(filePath, "utf8").catch(() => "Case study not found.");

  return (
    <main>
      <h1>{params.id}</h1>
      <pre>{content}</pre>
    </main>
  );
}
