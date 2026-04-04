declare global {

    interface StudentData {
        start_year: string,
        end_year: string,
        period: string,
        date: string,
        first_names: string,
        last_names: string,
        id: string,
        grade: string,
        level: string,
        absences?: number,
        p1_average?: number,
        p2_average?: number,
        p3_average?: number,
        pf_average?: number,
        groups: CourseGroup[],
        courses: CourseComment[],
        images: {
            sep: string,
            school: string,
            signature: string,
        }
        fonts: {
            regular: string,
            bold: string,
            italic: string,
        }
    }

    interface CourseGroup {
        name: string,
        color: string,
        courses: CourseGrades[],
    }

    interface CourseGrades {
        name: string,
        p1?: number,
        p2?: number,
        p3?: number,
        final?: number,
        h1: string,
        h2: string,
        h3: string,
        h4: string,
    }

    interface CourseComment {
        name: string,
        absences?: number,
        comment: string,
    }
}

export { };