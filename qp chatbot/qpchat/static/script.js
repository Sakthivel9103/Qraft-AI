const K_LEVEL_DEFS = {
    K1: "Remember",
    K2: "Understand",
    K3: "Apply",
    K4: "Analyze",
    K5: "Evaluate",
    K6: "Create",
};


const EXAM_DEFAULTS = {
    "CIA 1": {
        start: "K1",
        end: "K3",
        total_questions: 10,
        marks: 5
    },

    "CIA 2": {
        start: "K1",
        end: "K4",
        total_questions: 10,
        marks: 5
    },

    "Model Exam": {
        start: "K1",
        end: "K5",
        total_questions: 20,
        marks: 5
    },

    "Semester Exam": {
        start: "K1",
        end: "K6",
        total_questions: 20,
        marks: 5
    }
};


// =====================================================
// GET ELEMENTS
// =====================================================

const heroScreen = document.getElementById("heroScreen");
const chatScreen = document.getElementById("chatScreen");

const sidebarExamType = document.getElementById("sidebarExamType");

const examTitleInput = document.getElementById("exam_title");
const subjectInput = document.getElementById("subject");

const backToHomeBtn = document.getElementById("backToHomeBtn");

const settingsPanel = document.getElementById("settingsPanel");
const settingsToggle = document.getElementById("settingsToggle");

const generationMode = document.getElementById("generation_mode");

const normalSettings = document.getElementById("normalSettings");
const customSettings = document.getElementById("customSettings");

const marksInput = document.getElementById("marks_per_question");

const startLevelSelect = document.getElementById("start_level");
const endLevelSelect = document.getElementById("end_level");

const totalQuestionsInput =
    document.getElementById("total_questions");

const formatSelect =
    document.getElementById("file_format");

const settingsTotal =
    document.getElementById("settingsTotal");

const chatLog =
    document.getElementById("chatLog");

const sendBtn =
    document.getElementById("sendBtn");

const coreInput =
    document.getElementById("core_content");

const fileInput =
    document.getElementById("fileInput");

const attachBtn =
    document.getElementById("attachBtn");


// =====================================================
// CUSTOM PART INPUTS
// =====================================================

const partAQuestions =
    document.getElementById("partA_questions");

const partAMarks =
    document.getElementById("partA_marks");

const partAStart =
    document.getElementById("partA_start_level");

const partAEnd =
    document.getElementById("partA_end_level");


const partBQuestions =
    document.getElementById("partB_questions");

const partBMarks =
    document.getElementById("partB_marks");

const partBChoice =
    document.getElementById("partB_choice");

const partBStart =
    document.getElementById("partB_start_level");

const partBEnd =
    document.getElementById("partB_end_level");


const partCQuestions =
    document.getElementById("partC_questions");

const partCMarks =
    document.getElementById("partC_marks");

const partCStart =
    document.getElementById("partC_start_level");

const partCEnd =
    document.getElementById("partC_end_level");


// =====================================================
// BUILD K LEVEL OPTIONS
// =====================================================

function addKLevelOptions(select) {

    Object.entries(K_LEVEL_DEFS).forEach(([level, label]) => {

        const option =
            document.createElement("option");

        option.value = level;

        option.textContent =
            `${level} - ${label}`;

        select.appendChild(option);

    });

}


const allKLevelSelects = [

    startLevelSelect,
    endLevelSelect,

    partAStart,
    partAEnd,

    partBStart,
    partBEnd,

    partCStart,
    partCEnd

];


allKLevelSelects.forEach(select => {

    addKLevelOptions(select);

});


// =====================================================
// DEFAULT CUSTOM K LEVELS
// =====================================================

partAStart.value = "K1";
partAEnd.value = "K2";

partBStart.value = "K2";
partBEnd.value = "K4";

partCStart.value = "K3";
partCEnd.value = "K5";


// =====================================================
// GENERATION MODE CHANGE
// =====================================================

generationMode.addEventListener("change", () => {

    if (generationMode.value === "custom") {

        normalSettings.style.display = "none";

        customSettings.style.display = "block";

    } else {

        normalSettings.style.display = "block";

        customSettings.style.display = "none";

    }

    updateTotals();

});


// =====================================================
// UPDATE TOTALS
// =====================================================

function updateTotals() {

    if (generationMode.value === "normal") {

        const totalQ =
            parseInt(totalQuestionsInput.value || "0");

        const marks =
            parseInt(marksInput.value || "0");

        const totalMarks =
            totalQ * marks;

        settingsTotal.innerHTML =
            `<b>${totalQ}</b> questions ·
             <b>${totalMarks}</b> marks`;

    } else {

        const aQ =
            parseInt(partAQuestions.value || "0");

        const aM =
            parseInt(partAMarks.value || "0");


        const bQ =
            parseInt(partBQuestions.value || "0");

        const bM =
            parseInt(partBMarks.value || "0");


        const cQ =
            parseInt(partCQuestions.value || "0");

        const cM =
            parseInt(partCMarks.value || "0");


        const totalQuestions =
            aQ + bQ + cQ;

        const totalMarks =
            (aQ * aM) +
            (bQ * bM) +
            (cQ * cM);


        settingsTotal.innerHTML =
            `<b>${totalQuestions}</b> questions ·
             <b>${totalMarks}</b> marks`;

    }

}


// Update totals for all inputs
document.addEventListener("input", updateTotals);

document.addEventListener("change", updateTotals);


// =====================================================
// SETTINGS TOGGLE
// =====================================================

settingsToggle.addEventListener("click", () => {

    settingsPanel.classList.toggle("open");

});


// =====================================================
// AUTO RESIZE TEXTAREA
// =====================================================

coreInput.addEventListener("input", () => {

    coreInput.style.height = "auto";

    coreInput.style.height =
        Math.min(coreInput.scrollHeight, 160) + "px";

});


// =====================================================
// ENTER TO GENERATE
// =====================================================

coreInput.addEventListener("keydown", (e) => {

    if (e.key === "Enter" && !e.shiftKey) {

        e.preventDefault();

        sendBtn.click();

    }

});


// =====================================================
// FILE UPLOAD
// =====================================================

attachBtn.addEventListener("click", () => {

    fileInput.click();

});


fileInput.addEventListener("change", async () => {

    const file = fileInput.files[0];

    if (!file) return;


    addMsg(
        "bot",
        `Reading <b>${escapeHtml(file.name)}</b>...`
    );


    const formData = new FormData();

    formData.append("file", file);


    try {

        const res = await fetch(
            "/api/upload",
            {
                method: "POST",
                body: formData
            }
        );


        const data = await res.json();


        if (data.error) {

            addMsg(
                "bot",
                data.error,
                "msg--error"
            );

        } else {

            coreInput.value = data.text;

            coreInput.style.height = "auto";

            coreInput.style.height =
                Math.min(
                    coreInput.scrollHeight,
                    160
                ) + "px";


            addMsg(
                "bot",
                `Loaded <b>${escapeHtml(file.name)}</b>.
                 Review the syllabus and press Generate.`
            );

        }

    } catch (err) {

        addMsg(
            "bot",
            "Could not upload file: " +
            err.message,
            "msg--error"
        );

    }


    fileInput.value = "";

});


// =====================================================
// HERO TO CHAT
// =====================================================

document.querySelectorAll(".option-card")
    .forEach((btn) => {

        btn.addEventListener("click", () => {

            const examType =
                btn.dataset.exam;

            const preset =
                EXAM_DEFAULTS[examType];


            examTitleInput.value =
                examType;

            sidebarExamType.textContent =
                examType;


            marksInput.value =
                preset.marks;

            startLevelSelect.value =
                preset.start;

            endLevelSelect.value =
                preset.end;

            totalQuestionsInput.value =
                preset.total_questions;


            updateTotals();


            chatLog.innerHTML = "";


            addMsg(
                "bot",
                `Starting <b>${examType}</b>.
                Paste your syllabus or upload PDF/DOCX/TXT.
                You can use Normal Generation or Custom Question Paper Pattern from Settings ⚙.`
            );


            heroScreen.style.display =
                "none";

            chatScreen.style.display =
                "grid";


            coreInput.focus();

        });

    });


// =====================================================
// BACK TO HOME
// =====================================================

backToHomeBtn.addEventListener("click", () => {

    chatScreen.style.display =
        "none";

    heroScreen.style.display =
        "flex";

});


// =====================================================
// ADD CHAT MESSAGE
// =====================================================

function addMsg(
    role,
    html,
    extraClass = ""
) {

    const wrap =
        document.createElement("div");


    wrap.className =
        `msg msg--${role} ${extraClass}`;


    const avatar =
        role === "bot"

            ? `<div class="msg__avatar msg__avatar--bot">
                Q
               </div>`

            : `<div class="msg__avatar msg__avatar--user">
                You
               </div>`;


    wrap.innerHTML =
        `${avatar}
         <div class="msg__bubble">
            ${html}
         </div>`;


    chatLog.appendChild(wrap);

    chatLog.scrollTop =
        chatLog.scrollHeight;


    return wrap;

}


// =====================================================
// VALIDATE K LEVEL RANGE
// =====================================================

function isValidRange(start, end) {

    const levels =
        Object.keys(K_LEVEL_DEFS);

    return levels.indexOf(start) <=
           levels.indexOf(end);

}


// =====================================================
// GENERATE BUTTON
// =====================================================

sendBtn.addEventListener(
    "click",
    async () => {

        const core_content =
            coreInput.value.trim();

        const subject =
            subjectInput.value.trim();

        const exam_title =
            examTitleInput.value.trim();

        const format =
            formatSelect.value;


        if (!core_content) {

            addMsg(
                "bot",
                "Paste your syllabus first or upload a file.",
                "msg--error"
            );

            return;

        }


        const mode =
            generationMode.value;


        let endpoint;
        let payload;


        // =============================================
        // NORMAL MODE
        // =============================================

        if (mode === "normal") {

            const marks_per_question =
                marksInput.value;

            const start_level =
                startLevelSelect.value;

            const end_level =
                endLevelSelect.value;

            const total_questions =
                totalQuestionsInput.value;


            if (
                !isValidRange(
                    start_level,
                    end_level
                )
            ) {

                addMsg(
                    "bot",
                    "Invalid K-Level range.",
                    "msg--error"
                );

                return;

            }


            endpoint =
                "/api/generate";


            payload = {

                core_content,
                subject,
                exam_title,

                marks_per_question,

                start_level,
                end_level,

                total_questions,

                format

            };

        }


        // =============================================
        // CUSTOM PATTERN MODE
        // =============================================

        else {

            if (
                !isValidRange(
                    partAStart.value,
                    partAEnd.value
                )
            ) {

                addMsg(
                    "bot",
                    "PART A K-Level range is invalid.",
                    "msg--error"
                );

                return;

            }


            if (
                !isValidRange(
                    partBStart.value,
                    partBEnd.value
                )
            ) {

                addMsg(
                    "bot",
                    "PART B K-Level range is invalid.",
                    "msg--error"
                );

                return;

            }


            if (
                !isValidRange(
                    partCStart.value,
                    partCEnd.value
                )
            ) {

                addMsg(
                    "bot",
                    "PART C K-Level range is invalid.",
                    "msg--error"
                );

                return;

            }


            endpoint =
                "/api/generate-pattern";


            payload = {

                core_content,
                subject,
                exam_title,
                format,

                pattern: [

                    {

                        section: "PART A",

                        question_count:
                            parseInt(
                                partAQuestions.value
                            ),

                        marks:
                            parseInt(
                                partAMarks.value
                            ),

                        has_choice:
                            false,

                        start_level:
                            partAStart.value,

                        end_level:
                            partAEnd.value

                    },


                    {

                        section: "PART B",

                        question_count:
                            parseInt(
                                partBQuestions.value
                            ),

                        marks:
                            parseInt(
                                partBMarks.value
                            ),

                        has_choice:
                            partBChoice.value ===
                            "yes",

                        start_level:
                            partBStart.value,

                        end_level:
                            partBEnd.value

                    },


                    {

                        section: "PART C",

                        question_count:
                            parseInt(
                                partCQuestions.value
                            ),

                        marks:
                            parseInt(
                                partCMarks.value
                            ),

                        has_choice:
                            false,

                        start_level:
                            partCStart.value,

                        end_level:
                            partCEnd.value

                    }

                ]

            };

        }


        // SHOW USER MESSAGE

        addMsg(
            "user",
            escapeHtml(core_content).slice(0, 1000)
        );


        coreInput.value = "";

        coreInput.style.height = "auto";


        sendBtn.disabled = true;


        const loadingMsg =
            addMsg(
                "bot",
                "Drafting your question paper..."
            );


        try {

            const res =
                await fetch(
                    endpoint,
                    {

                        method: "POST",

                        headers: {
                            "Content-Type":
                                "application/json"
                        },

                        body:
                            JSON.stringify(payload)

                    }
                );


            const data =
                await res.json();


            loadingMsg.remove();


            if (data.error) {

                addMsg(
                    "bot",
                    data.error,
                    "msg--error"
                );

            } else {

                const formatLabel =
                    format === "pdf"

                        ? "question_paper.pdf"

                        : "question_paper.docx";


                addMsg(
                    "bot",
                    `

                    <p style="margin:0 0 10px;">

                    <b>${data.total_questions}</b>
                    questions ·

                    <b>${data.total_marks}</b>
                    marks total.

                    </p>


                    <div class="paper-preview">

                    ${escapeHtml(
                        data.paper_text
                    )}

                    </div>


                    <div class="download-row">

                    <span>
                    ${formatLabel} ready
                    </span>


                    <a
                        href="${data.download_url}"
                        download>

                        Download

                    </a>

                    </div>

                    `
                );

            }

        }

        catch (err) {

            loadingMsg.remove();


            addMsg(
                "bot",
                "Something went wrong: " +
                err.message,
                "msg--error"
            );

        }

        finally {

            sendBtn.disabled = false;

        }

    }
);


// =====================================================
// ESCAPE HTML
// =====================================================
function escapeHtml(str) {
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\n/g, "<br>");
}
// =====================================================
// INITIAL TOTAL
// =====================================================
updateTotals();