export const openedByOptions = [
  { name: "Md. Hasanuzzaman", designation: "Head Of Business Expansion" },
  { name: "Mr. Shamsuzzaman", designation: "Assistant Director, Business Expansion" },
  { name: "Mr. Mohammad Zaherul Islam", designation: "Assistant Director, Business Expansion" },
  { name: "Mr. Abdullah Al Masum", designation: "Head Of Business Expansion" },
  { name: "Mr. Lutful Kabir Munna", designation: "Head Of Business Expansion" },
  { name: "Mr. Fakhrul Alamgir", designation: "Head Of Business Expansion" },
  { name: "Mr. Abu Sayem Mohammad Al Azad", designation: "Head Of Business Expansion" },
  { name: "Mr. Kaushik", designation: "Regional Head" },
  { name: "Mr. Shadhin", designation: "Regional Head" },
  { name: "Mr. Azim", designation: "Regional Head" },
  { name: "Mr. Emran", designation: "Regional Head" },
  { name: "Mr. Mahbub", designation: "Regional Head" },
  { name: "Mr. Kabir", designation: "Regional Head" },
  { name: "Mr. Riaz", designation: "Regional Head" },
  { name: "Mr. Sunny", designation: "Regional Head" },
  { name: "Mr. Yusuf", designation: "Regional Head" },
  { name: "Ms. Qarin", designation: "Regional Head" },
  { name: "Mr. Ranjan", designation: "Regional Head" },
  { name: "Mr. Mehedi Hasan Limon", designation: "Head Of Franchisee Relations" },
  { name: "Ferdous Zaman", designation: "Manager Franchise Expansion" },
];

// Opened by names added after the workbook's Master!E8:F15 lookup table was
// made. The Excel with Rules writes them below that table (from Master!E16)
// so Master!C17 =VLOOKUP(C16,...) also returns their designation in Excel.
export const OPENED_BY_MASTER_LOOKUP_ADDITIONS = Object.freeze(["Mr. Mehedi Hasan Limon", "Ferdous Zaman"]);

export const salesGivenByOptions = [
  "Mr. Mamun",
  "Anamul Hoq",
  "MD.MEHEDI HASAN ANTOR",
  "Mohammad Afzal Hossain Bhuiyan",
  "Bipul Kumar Das",
  "Imran Mahmud",
  "Iqbal Mahbub",
  "M.A Kayum Anik",
  "Md. Ershadul Haque",
  "Md. Fazlul Karim Prince",
  "Md. Jahid Hasan Khan",
  "Md. Jahidul Islam",
  "Md. Mamun Ahmed Munna",
  "Md. Shajjad Hossain Chowdhury",
  "Omar Faruk",
  "Sheikh Ashraf",
  "Abdur Razzak Akon",
  "Hasan Ahmed",
  "Md. Aminul Islam",
  "Md. Golam Kibria Khan",
  "Sajjad Hossain Sumon",
  "Subrata Ghose",
  "Md. Abdul Alim",
  "Md. Abu Musa Al-Tarique",
  "Md. Nurunnabi Islam (Choyon)",
  "Md. Shahinoor Alam",
  "Md. Moniruzzaman",
  "Niyaj Mahmud",
  "Akram Uz Zaman",
  "Jiku Miah",
  "Md Homaun Kabir Chowdhury Saju",
  "Md. Safiqul Islam",
  "MD Fardin Hossain",
  "Md Toriqul Islam",
  "Shahariar Sumon",
  "Easin Mia",
  "Md Faysal Haque",
  "Riad Mahmud Polash",
  "Sourav Paul Mithun",
  "Tanzin Hossain",
  "Kamrul Hasan",
  "Md. Mainul Islam",
  "Md. Rashel",
  "Md. Abdul Azad",
  "Mohammad Year Hossain",
  "Siful Islam Maruf",
  "Atiqul Islam",
  "Hafiz Mohammed Iqbal",
  "Ibrahim Khalilullah",
  "Jakir Hossain",
  "Md Khairul Islam",
  "Md Omar Faruque",
  "Md Samrat Miah",
  "Md. Eazul Islam (Mahin)",
  "Md. Mijanur Rahman",
  "Md. Mohidul Islam (Herok)",
  "Md. Mohiuddin",
  "Md. Rasel Al Mamun",
  "Md. Shahidul Islam",
  "Nahid Hasan",
];

export const districtDivisionOptions = Object.freeze([
  { district: "Bagerhat", division: "Khulna" },
  { district: "Bandarban", division: "Chattogram" },
  { district: "Barguna", division: "Barishal" },
  { district: "Barishal", division: "Barishal" },
  { district: "Bhola", division: "Barishal" },
  { district: "Bogura", division: "Rajshahi" },
  { district: "Brahmanbaria", division: "Chattogram" },
  { district: "Chandpur", division: "Chattogram" },
  { district: "Chapainawabganj", division: "Rajshahi" },
  { district: "Chattogram", division: "Chattogram" },
  { district: "Chuadanga", division: "Khulna" },
  { district: "Cox's Bazar", division: "Chattogram" },
  { district: "Cumilla", division: "Chattogram" },
  { district: "Dhaka", division: "Dhaka" },
  { district: "Dhaka", division: "DhakaGBUD" },
  { district: "Dinajpur", division: "Rangpur" },
  { district: "Faridpur", division: "Dhaka" },
  { district: "Feni", division: "Chattogram" },
  { district: "Gaibandha", division: "Rangpur" },
  { district: "Gazipur", division: "Dhaka" },
  { district: "Gopalganj", division: "Dhaka" },
  { district: "Habiganj", division: "Sylhet" },
  { district: "Jamalpur", division: "Mymensingh" },
  { district: "Jashore", division: "Khulna" },
  { district: "Jhalokathi", division: "Barishal" },
  { district: "Jhenaidah", division: "Khulna" },
  { district: "Joypurhat", division: "Rajshahi" },
  { district: "Khagrachhari", division: "Chattogram" },
  { district: "Khulna", division: "Khulna" },
  { district: "Kishoreganj", division: "Mymensingh" },
  { district: "Kurigram", division: "Rangpur" },
  { district: "Kushtia", division: "Khulna" },
  { district: "Lakshmipur", division: "Chattogram" },
  { district: "Lalmonirhat", division: "Rangpur" },
  { district: "Madaripur", division: "Dhaka" },
  { district: "Magura", division: "Khulna" },
  { district: "Manikganj", division: "Dhaka" },
  { district: "Meherpur", division: "Khulna" },
  { district: "Moulvibazar", division: "Sylhet" },
  { district: "Munshiganj", division: "Dhaka" },
  { district: "Mymensingh", division: "Mymensingh" },
  { district: "Naogaon", division: "Rajshahi" },
  { district: "Narail", division: "Khulna" },
  { district: "Narayanganj", division: "Dhaka" },
  { district: "Narsingdi", division: "Dhaka" },
  { district: "Natore", division: "Rajshahi" },
  { district: "Netrokona", division: "Mymensingh" },
  { district: "Nilphamari", division: "Rangpur" },
  { district: "Noakhali", division: "Chattogram" },
  { district: "Pabna", division: "Rajshahi" },
  { district: "Panchagarh", division: "Rangpur" },
  { district: "Patuakhali", division: "Barishal" },
  { district: "Pirojpur", division: "Barishal" },
  { district: "Rajbari", division: "Dhaka" },
  { district: "Rajshahi", division: "Rajshahi" },
  { district: "Rangamati", division: "Chattogram" },
  { district: "Rangpur", division: "Rangpur" },
  { district: "Satkhira", division: "Khulna" },
  { district: "Savar", division: "Dhaka" },
  { district: "Shariatpur", division: "Dhaka" },
  { district: "Sherpur", division: "Mymensingh" },
  { district: "Sirajganj", division: "Rajshahi" },
  { district: "Sunamganj", division: "Sylhet" },
  { district: "Sylhet", division: "Sylhet" },
  { district: "Tangail", division: "Dhaka" },
  { district: "Thakurgaon", division: "Rangpur" },
]);
export const DEFAULT_SIGNATORIES = [
    {
      role: "Prepared by",
      name: "Md Aftab Ul Islam",
      designation: "Senior Executive, Retail Operations",
      signatureId: "source-signature-2",
      includeInPdf: true,
    },
    {
      role: "Sales Given By",
      name: "Mr. Mamun",
      designation: "Leader's Name",
      autoSource: "sales-given-by",
      manualOverride: false,
      signatureId: "source-signature-1",
      includeInPdf: true,
    },
    {
      role: "Opened by",
      name: "Mr. Abu Sayem Mohammad Al Azad",
      designation: "Head Of Business Expansion",
      autoSource: "opened-by",
      manualOverride: false,
      signatureId: "",
      includeInPdf: false,
    },
    {
      role: "Viewed by",
      name: "Salah Uddin Misbah",
      designation: "Head Of Business, CG",
      signatureId: "salah-uddin-misbah-signature",
      // Available from Signature Manager, but not selected in a new PDF.
      includeInPdf: false,
    },
    {
      role: "Viewed by",
      name: "Saiful Alam Rasel",
      designation: "Assistant Director, Operations",
      signatureId: "saiful-alam-rasel-signature",
      // Included by default and also used as the review signature on pages 1 and 2.
      includeInPdf: true,
    },
    {
      role: "Viewed by",
      name: "Tamal Paul",
      designation: "Director – Business Planning & Investor Relations",
      signatureId: "tamal-paul-signature",
      includeInPdf: false,
    },
    {
      role: "Approved By",
      name: "Abu Naser",
      designation: "Chief Business Officer",
      signatureId: "",
      includeInPdf: false,
    },
  ];
export function getSignatoryAutoLink(signatory) {
  const source = Object.hasOwn(signatory || {}, "autoSource")
    ? signatory.autoSource
    : signatory?.role === "Sales Given By"
      ? "sales-given-by"
      : signatory?.role === "Opened by"
        ? "opened-by"
        : null;

  if (source === "sales-given-by") {
    return {
      source,
      label: "Sales Given By",
      updatesDesignation: false,
    };
  }
  if (source === "opened-by") {
    return {
      source,
      label: "Opened by",
      updatesDesignation: true,
    };
  }
  return null;
}


export function resolveSignatories(data) {
 const project=data.project;
  const signatories = data.signatories.map((signatory) => {
    const autoLink = getSignatoryAutoLink(signatory);
    if (!autoLink || signatory.manualOverride === true) return { ...signatory };
    if (autoLink.source === "sales-given-by") return { ...signatory, name: project.salesGivenBy };
    if (autoLink.source === "opened-by") return {
      ...signatory,
      name: project.openedBy,
      designation: project.openedBy ? (project.openedDesignation || signatory.designation) : "",
    };
    return { ...signatory };
  });

  // Signatory 3 must never carry its own independent signature. When its PDF
  // checkbox is enabled, it uses the exact signature currently selected for
  // Signatory 2; when disabled, it has no signature in either export preview.
  // Find them by their source tags so the relationship still works if other
  // signatories are added or their display roles are edited.
  const signatory2Index = data.signatories.findIndex((signatory) => (
    getSignatoryAutoLink(signatory)?.source === "sales-given-by"
  ));
  const signatory3Index = data.signatories.findIndex((signatory) => (
    getSignatoryAutoLink(signatory)?.source === "opened-by"
  ));
  if (signatory3Index >= 0) {
    const signatory3 = signatories[signatory3Index];
    const signatory2 = signatory2Index >= 0 ? signatories[signatory2Index] : null;
    signatories[signatory3Index] = {
      ...signatory3,
      signatureId: signatory3.includeInPdf === true ? (signatory2?.signatureId || "") : "",
    };
  }

return signatories;
}
export function getOpenedDesignation(name) {return openedByOptions.find(p=>p.name===name)?.designation ?? "";}
