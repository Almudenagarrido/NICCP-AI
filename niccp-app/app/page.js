import fs from "fs";
import path from "path";
import Link from "next/link";

// Esto va a ejecutarse en el servidor, así que fs funciona
export default function Home() {
  const caseStudiesDir = path.join(process.cwd(), "app", "case-studies");
  const templatePath = path.join(caseStudiesDir, "case-template.json");
  let files = fs.readdirSync(caseStudiesDir).filter(
    (file) => file.endsWith(".json") && file !== "case-template.json"
  );

  async function createNewCase(formData) {
    "use server";

    const name = formData.get("name");
    const description = formData.get("description");

    const template = JSON.parse(fs.readFileSync(templatePath, "utf-8"));
    template["general-data"] = { name, description };

    const newFilePath = path.join(caseStudiesDir, `${name}.json`);
    fs.writeFileSync(newFilePath, JSON.stringify(template, null, 2));
  }

  return (
    <>
      <h1>Manage case studies</h1>
      <ul>
        {files.map((file) => {
          const id = file.replace(".json", "");
          return (
            <li key={file}>
              <Link href={`/case-studies/${id}/general-data`}>{id}</Link>
            </li>
          );
        })}
      </ul>

      <form action={createNewCase}>
        <h4>Create new case study</h4>
        <div>
          <input name="name" type="text" placeholder="Name" required />
        </div>
        <div>
          <input name="description" type="text" placeholder="Description" required />
        </div>
        <button type="submit">Create</button>
      </form>
    </>
  );
}
