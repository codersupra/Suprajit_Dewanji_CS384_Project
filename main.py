import pandas as pd
from collections import defaultdict

# Load input Excel files
students_data = pd.read_excel('ip_1.xlsx')
schedule_data = pd.read_excel('ip_2.xlsx')
rooms_data = pd.read_excel('ip_3.xlsx')

# User inputs
arrangement_mode = int(input("Press 1 for sparse arrangement or 2 for dense arrangement: "))
buffer_space = int(input("Enter buffer size per classroom (default is 5): ") or 5)

# Calculate student count per course
student_counts = students_data['course_code'].value_counts().to_dict()

# Sort rooms by block and capacity
block_9_rooms = rooms_data[rooms_data['Block'] == 9].sort_values(by='Exam Capacity', ascending=False)
lt_rooms = rooms_data[rooms_data['Block'] == 'LT'].sort_values(by='Exam Capacity', ascending=False)

# Parse exam schedule
exam_schedule = defaultdict(lambda: {'Morning': [], 'Evening': []})
for _, row in schedule_data.iterrows():
    exam_schedule[row['Date']] = {
        'Morning': row['Morning'].split('; ') if row['Morning'] != "NO EXAM" else [],
        'Evening': row['Evening'].split('; ') if row['Evening'] != "NO EXAM" else []
    }

# Generate seating plan and room summary
seating_plan = []
room_summary = []

for date, sessions in exam_schedule.items():
    for session, courses in sessions.items():
        sorted_courses = sorted(courses, key=lambda x: student_counts.get(x, 0), reverse=True)
        
        for course in sorted_courses:
            enrolled_students = students_data[students_data['course_code'] == course]['rollno'].tolist()
            assigned_index = 0
            total_students = len(enrolled_students)

            for room_set in [block_9_rooms, lt_rooms]:
                for _, room in room_set.iterrows():
                    available_capacity = room['Exam Capacity'] - buffer_space
                    max_seats = available_capacity // 2 if arrangement_mode == 1 else available_capacity
                    assigned_count = min(total_students - assigned_index, max_seats)

                    if assigned_count > 0:
                        assigned_rolls = ";".join(enrolled_students[assigned_index:assigned_index + assigned_count])
                        seating_plan.append([date, session, course, room['Room No.'], assigned_count, assigned_rolls])
                        assigned_index += assigned_count

                    if assigned_index >= total_students:
                        break
                if assigned_index >= total_students:
                    break

for _, room in rooms_data.iterrows():
    room_no = room['Room No.']
    capacity = room['Exam Capacity']
    occupied_seats = sum(row[4] for row in seating_plan if row[3] == room_no) + buffer_space
    room_summary.append([room_no, capacity, room['Block'], max(0, capacity - occupied_seats)])

# Save output to Excel
seating_plan_df = pd.DataFrame(seating_plan, columns=['Date', 'Session', 'Course', 'Room', 'Students Count', 'Roll List'])
room_summary_df = pd.DataFrame(room_summary, columns=['Room No.', 'Capacity', 'Block', 'Vacant Seats'])

with pd.ExcelWriter('output.xlsx') as writer:
    seating_plan_df.to_excel(writer, sheet_name='Seating Plan', index=False)
    room_summary_df.to_excel(writer, sheet_name='Room Summary', index=False)

print("Output saved to 'output.xlsx'.")
