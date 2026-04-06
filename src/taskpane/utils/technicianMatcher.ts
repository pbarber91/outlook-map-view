type TechnicianRecord = {
  name: string;
  email: string;
};

const TECHNICIANS: TechnicianRecord[] = [
  { name: "Joey Schilk", email: "joey@ffe-fl.com" },
  { name: "Jeremy Whitaker", email: "jeremy@ffe-fl.com" },
  { name: "Anthony Silva", email: "anthony@ffe-fl.com" },
  { name: "Bill Fischer", email: "williamfischer@ffe-fl.com" },
  { name: "Garrett Borth", email: "garrett@ffe-fl.com" },
  { name: "Nick Van Buskirk", email: "nick@ffe-fl.com" },
  { name: "Nolan Gora", email: "nolan@ffe-fl.com" },
  { name: "Austin Rinke", email: "austin@ffe-fl.com" },
];

function normalize(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^a-z0-9@.\s]/g, "");
}

function normalizeNameForLooseMatch(value: string): string {
  return normalize(value).replace(/\s+/g, "");
}

const TECHNICIAN_INDEX = TECHNICIANS.map((tech) => ({
  canonicalName: tech.name,
  normalizedName: normalize(tech.name),
  normalizedNameLoose: normalizeNameForLooseMatch(tech.name),
  normalizedEmail: normalize(tech.email),
  emailLocalPart: normalize(tech.email.split("@")[0] || ""),
}));

export function matchTechnician(
  displayName?: string | null,
  email?: string | null
): string | null {
  const normalizedDisplayName = displayName ? normalize(displayName) : "";
  const normalizedDisplayNameLoose = displayName
    ? normalizeNameForLooseMatch(displayName)
    : "";
  const normalizedEmail = email ? normalize(email) : "";
  const normalizedEmailLocalPart = normalizedEmail.includes("@")
    ? normalizedEmail.split("@")[0]
    : normalizedEmail;

  for (const tech of TECHNICIAN_INDEX) {
    if (normalizedEmail && normalizedEmail === tech.normalizedEmail) {
      return tech.canonicalName;
    }

    if (normalizedDisplayName && normalizedDisplayName === tech.normalizedName) {
      return tech.canonicalName;
    }

    if (
      normalizedDisplayNameLoose &&
      normalizedDisplayNameLoose === tech.normalizedNameLoose
    ) {
      return tech.canonicalName;
    }

    if (
      normalizedDisplayName &&
      (normalizedDisplayName === tech.emailLocalPart ||
        normalizedDisplayNameLoose === tech.emailLocalPart)
    ) {
      return tech.canonicalName;
    }

    if (
      normalizedEmailLocalPart &&
      normalizedEmailLocalPart === tech.emailLocalPart
    ) {
      return tech.canonicalName;
    }
  }

  return null;
}

export function getWhitelistedTechnicians(): string[] {
  return TECHNICIANS.map((tech) => tech.name);
}