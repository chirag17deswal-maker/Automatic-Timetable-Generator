
document.addEventListener("DOMContentLoaded", function () {
    const menuButton = document.getElementById("menuToggle");
    const navLinks = document.querySelector(".nav-links");

    if (menuButton && navLinks) {
        menuButton.addEventListener("click", function () {
            const isOpen = navLinks.classList.toggle("open");

            menuButton.setAttribute(
                "aria-expanded",
                String(isOpen)
            );
        });

        navLinks.querySelectorAll("a").forEach(function (link) {
            link.addEventListener("click", function () {
                navLinks.classList.remove("open");
                menuButton.setAttribute("aria-expanded", "false");
            });
        });
    }
});
// ================================
// Academic Structure
// ================================

const addCourseBtn = document.getElementById("addCourseBtn");
const coursesContainer = document.getElementById("coursesContainer");

let courseCount = 0;


if (addCourseBtn && coursesContainer) {
    // Add Course
    addCourseBtn.addEventListener("click", () => {

    courseCount++;

    const course = document.createElement("div");

    course.className = "course-card";

    course.innerHTML = `
        <div class="course-header">

            <div>
                <span class="item-number">
                    COURSE ${courseCount}
                </span>

                <h3>New Course</h3>
            </div>

            <button type="button" class="remove-btn">
                Remove
            </button>

        </div>

        <div class="form-group">

            <label>Course Name</label>

            <input
                type="text"
                class="course-name"
                placeholder="e.g. B.Tech, BCA, BBA, Diploma"
            >

        </div>

        <div class="branches-section">

            <div class="subsection-header">

                <div>
                    <h4>Branches / Programs</h4>
                    <p>Add branches under this course.</p>
                </div>

                <button type="button" class="add-branch-btn">
                    + Add Branch
                </button>

            </div>

            <div class="branches-container"></div>

        </div>
    `;


    coursesContainer.appendChild(course);


    // Remove Course
    const removeBtn = course.querySelector(".remove-btn");

    removeBtn.addEventListener("click", () => {
        course.remove();
    });


    // Add Branch
    const addBranchBtn = course.querySelector(".add-branch-btn");

    const branchesContainer =
        course.querySelector(".branches-container");

    let branchCount = 0;


    addBranchBtn.addEventListener("click", () => {

        branchCount++;

        const branch = document.createElement("div");

        branch.className = "branch-card";

        branch.innerHTML = `

            <div class="branch-header">

                <input
                    type="text"
                    class="branch-name"
                    placeholder="e.g. Computer Science"
                >

                <button type="button" class="remove-branch-btn">
                    Remove
                </button>

            </div>

            <div class="semester-section">

                <div class="subsection-header">

                    <div>
                        <h4>Semesters</h4>
                        <p>Add semesters for this branch.</p>
                    </div>

                    <button type="button" class="add-semester-btn">
                        + Add Semester
                    </button>

                </div>

                <div class="semesters-container"></div>

            </div>
        `;


        branchesContainer.appendChild(branch);


        // Remove Branch
        const removeBranchBtn =
            branch.querySelector(".remove-branch-btn");

        removeBranchBtn.addEventListener("click", () => {
            branch.remove();
        });


        // Add Semester
        const addSemesterBtn =
            branch.querySelector(".add-semester-btn");

        const semestersContainer =
            branch.querySelector(".semesters-container");


        addSemesterBtn.addEventListener("click", () => {

            const semester = document.createElement("div");

            semester.className = "semester-card";

            semester.innerHTML = `

                <div class="semester-header">

                    <input
                        type="text"
                        class="semester-name"
                        placeholder="e.g. Semester 1"
                    >

                    <button
                        type="button"
                        class="remove-semester-btn">
                        Remove
                    </button>

                </div>

                <div class="section-area">

                    <div class="subsection-header">

                        <div>
                            <h4>Sections</h4>
                            <p>
                                Add sections for this semester.
                            </p>
                        </div>

                        <button
                            type="button"
                            class="add-section-btn">
                            + Add Section
                        </button>

                    </div>

                    <div class="sections-container"></div>

                </div>
            `;


            semestersContainer.appendChild(semester);


            // Remove Semester
            semester
                .querySelector(".remove-semester-btn")
                .addEventListener("click", () => {
                    semester.remove();
                });


            // Add Section
            const addSectionBtn =
                semester.querySelector(".add-section-btn");

            const sectionsContainer =
                semester.querySelector(".sections-container");


            addSectionBtn.addEventListener("click", () => {

                const section = document.createElement("div");

                section.className = "section-item";

                section.innerHTML = `

                    <input
                        type="text"
                        class="section-name"
                        placeholder="Section A"
                    >

                    <button
                        type="button"
                        class="remove-section-btn">
                        ×
                    </button>

                `;


                sectionsContainer.appendChild(section);


                section
                    .querySelector(".remove-section-btn")
                    .addEventListener("click", () => {
                        section.remove();
                    });

            });

        });

    });

    });
}

const continueBtn = document.getElementById("continueBtn");
if (continueBtn) {
    continueBtn.addEventListener("click", () => {
        console.log("Save & Continue clicked");
    });
}