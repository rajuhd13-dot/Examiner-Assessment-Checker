// src/fallbackData.ts

// Dynamic mock generator for initial client states and fault-tolerant server-side fallbacks
export const generateFallbackRows = (): any[][] => {
  const institutions = [
    "Dhaka University (DU)",
    "Bangladesh University of Engineering and Technology (BUET)",
    "Jahangirnagar University (JU)",
    "Rajshahi University (RU)",
    "Chittagong University (CU)",
    "Khulna University (KU)",
    "Shahjalal University of Science and Technology (SUST)",
    "North South University (NSU)",
    "Brac University",
    "Mymensingh Medical College",
    "Dhaka Medical College"
  ];

  const departments = [
    "Computer Science and Engineering (CSE)",
    "Electrical and Electronic Engineering (EEE)",
    "Mechanical Engineering (ME)",
    "Physics",
    "Chemistry",
    "Mathematics",
    "English",
    "Bangla",
    "Applied Physics",
    "Statistics",
    "Microbiology"
  ];

  const names = [
    "Sadia Rahman", "Tariqul Islam", "Anika Tabassum", "Sabbir Rahman", "Afsana Mimi",
    "Fahim Miah", "Tanvir Ahmed", "Kazi Sakib", "Mehedi Hasan", "Nusrat Jahan",
    "Rashedul Islam", "Kamrul Hasan", "Imran Khan", "Zarin Tasnim", "Ahsan Kabir",
    "Taskin Ahmed", "Sabina Yasmin", "Asif Iqbal", "Moshiur Rahman", "Farhana Islam",
    "Syed Al-Amin", "Mahmudul Hasan", "Rubaiya Yeasmin", "Ariful Islam", "Tasnim Sultana"
  ];

  const statuses = ["Completed", "Pending", "Failed Evaluation", "Under Review", "Awaiting Date"];
  const districts = ["Dhaka", "Chittagong", "Sylhet", "Rajshahi", "Khulna", "Barisal", "Rangpur", "Mymensingh", "Comilla", "Gazipur", "Narayanganj"];
  const campuses = ["Farmgate", "Mirpur", "Uttara", "Mouchak", "Jatrabari", "Chittagong Centre", "Rajshahi Centre"];

  const rows: any[][] = [];

  // Always seed one specific ID the user/admin might type, for perfect demo validations
  // e.g. T2308129 (mentioned in previous context)
  const seedRow: any[] = Array(100).fill('');
  seedRow[0] = "1";                         // SL
  seedRow[1] = "Mirza Hassan";              // Nick Name
  seedRow[3] = "T2308129";                  // T-PIN
  seedRow[4] = "BUET";                      // Institution
  seedRow[5] = "Mechanical Engineering";    // Department
  seedRow[6] = "48";                        // HSC Batch
  seedRow[7] = "RM 03";                     // RM
  seedRow[8] = "System Administrator";      // Remarked By
  seedRow[9] = "01712345678";               // Mobile
  seedRow[10] = "01912345678";              // Alternate Mobile
  seedRow[11] = "01712345678";              // Nagad
  seedRow[16] = "Engineering Prep";         // Running Program
  seedRow[22] = "mirza.hassan@buet.edu";    // Email
  seedRow[30] = "Dhaka Board";              // Board
  seedRow[31] = "5.00";                     // HSC GPA
  seedRow[34] = "Physics";                  // Choice 1
  seedRow[35] = "Chemistry";                // Choice 2
  seedRow[36] = "Math";                     // Choice 3
  seedRow[61] = "35/100";                   // English (Fail)
  seedRow[64] = "60/100";                   // Bangla (Pass)
  seedRow[67] = "75/100";                   // Physics (Pass)
  seedRow[70] = "82/100";                   // Chemistry (Pass)
  seedRow[73] = "90/100";                   // Math (Pass)
  seedRow[76] = "55/100";                   // Biology (Pass)
  seedRow[79] = "85/100";                   // ICT (Pass)
  seedRow[82] = "Completed";                // Training Report
  seedRow[83] = "2026-04-10";               // Training Date
  seedRow[88] = "Farmgate";                 // Campus
  seedRow[92] = "Highly attentive and recommended for advanced tutoring loops."; // Remark Raw
  rows.push(seedRow);

  // Generate 54 more high fidelity records
  for (let i = 2; i <= 55; i++) {
    const r: any[] = Array(100).fill('');
    const randomName = names[i % names.length] + " " + String.fromCharCode(65 + (i % 26));
    const randomTpin = "T" + String(2300000 + i * 1489);
    const randomPhone = `01${Math.floor(3 + Math.random() * 7)}${Math.floor(10000000 + Math.random() * 90000000)}`;
    const randomAlt = `01${Math.floor(3 + Math.random() * 7)}${Math.floor(10000000 + Math.random() * 90000000)}`;
    const batch = String(40 + (i % 12));
    const rmNum = i % 8;
    const rmStr = rmNum === 0 ? "—" : `RM 0${rmNum}`;

    r[0] = String(i);
    r[1] = randomName;
    r[3] = randomTpin;
    r[4] = institutions[i % institutions.length];
    r[5] = departments[i % departments.length];
    r[6] = batch;
    r[7] = rmStr;
    r[8] = rmNum > 0 && rmNum % 3 === 0 ? "Sujon K." : "";
    r[9] = randomPhone;
    r[10] = randomAlt;
    r[11] = randomPhone;
    r[16] = i % 2 === 0 ? "Medical Prep" : "Varsity Prep";
    r[22] = `${randomName.toLowerCase().replace(/\s/g, '.')}@gmail.com`;
    r[30] = districts[i % districts.length] + " Board";
    r[31] = (4.50 + (i % 6) * 0.1).toFixed(2);
    r[34] = "English";
    r[35] = "Bangla";
    r[36] = i % 3 === 0 ? "ICT" : "Physics";
    
    // Custom scores 
    r[61] = `${Math.floor(30 + (i * 3) % 70)}/100`; // English
    r[64] = `${Math.floor(40 + (i * 4) % 60)}/100`; // Bangla
    r[67] = `${Math.floor(35 + (i * 5) % 65)}/100`; // Physics
    r[70] = `${Math.floor(40 + (i * 2) % 60)}/100`; // Chemistry
    r[73] = `${Math.floor(25 + (i * 7) % 75)}/100`; // Math
    r[76] = `${Math.floor(45 + (i * i) % 55)}/100`; // Biology
    r[79] = `${Math.floor(40 + (i * 6) % 60)}/100`; // ICT

    r[82] = statuses[i % statuses.length];
    r[83] = `2026-05-${String(1 + (i % 25)).padStart(2, '0')}`;
    r[88] = campuses[i % campuses.length];
    r[92] = i % 5 === 0 ? "Requires performance monitor check-up." : "";

    rows.push(r);
  }

  return rows;
};

export const FALLBACK_THRESHOLD_LIMITS = {
  english: 55,
  bangla: 48,
  physics: 48,
  chemistry: 48,
  math: 48,
  biology: 48,
  ict: 48
};
