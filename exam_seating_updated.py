import pandas as pd
import os
import re
import logging
from collections import defaultdict
import xlsxwriter
import math

# Configure logging
logging.basicConfig(
    filename="exam_conflicts.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def get_user_inputs():
    """Get user inputs for buffer, sparse/dense allocation preferences"""
    print("\n=== EXAM SEATING CONFIGURATION ===")
    
    # Get buffer value
    while True:
        try:
            buffer = int(input("Enter buffer value (number of seats to keep empty per room): "))
            if buffer >= 0:
                break
            else:
                print("Buffer value must be non-negative. Please try again.")
        except ValueError:
            print("Please enter a valid integer for buffer value.")
    
    # Get allocation preference
    print("\nAllocation Options:")
    print("0 = SPARSE (50% capacity, allow multiple subjects per room)")
    print("1 = DENSE (100% capacity, one subject per room)")
    
    while True:
        try:
            allocation_type = int(input("Select allocation type (0 for SPARSE, 1 for DENSE): "))
            if allocation_type in [0, 1]:
                break
            else:
                print("Please enter 0 for SPARSE or 1 for DENSE.")
        except ValueError:
            print("Please enter a valid integer (0 or 1).")
    
    sparse_factor = 0.5 if allocation_type == 0 else 1.0
    is_sparse = allocation_type == 0
    
    print(f"\n✅ Configuration Set:")
    print(f"   Buffer: {buffer} seats per room")
    print(f"   Allocation: {'SPARSE (50%)' if is_sparse else 'DENSE (100%)'}")
    
    return buffer, sparse_factor, is_sparse

def clean_timetable(cell):
    """Clean timetable cell data"""
    if isinstance(cell, str):
        parts = [x.strip() for x in cell.split(';') if x.strip()]
        return ';'.join(parts) if parts else None
    return cell

def generate_subjectwise_student_count(df_timetable, df_course_roll_mapping, output_file="subjectwise_student_count_sorted.xlsx"):
    """Generate Excel file containing student counts per subject per exam date"""
    try:
        # Extract (Date, SubjectCode) pairs
        subject_date_list = []
        for _, row in df_timetable.iterrows():
            date = row['Date']
            for col in ['Morning', 'Evening']:
                if pd.notnull(row[col]) and row[col] != "NO EXAM":
                    subjects = [s.strip() for s in str(row[col]).split(';') if s.strip()]
                    for subject in subjects:
                        subject_date_list.append((date, subject))

        df_subject_date = pd.DataFrame(subject_date_list, columns=['Date', 'SubjectCode'])

        # Merge with student-course mapping
        merged_df = df_subject_date.merge(
            df_course_roll_mapping,
            left_on='SubjectCode',
            right_on='course_code',
            how='left'
        )

        # Count unique students per subject per date
        count_df = merged_df.groupby(['Date', 'SubjectCode'])['rollno'].nunique().reset_index()
        count_df.rename(columns={'rollno': 'StudentCount'}, inplace=True)
        count_df = count_df.sort_values(by=['Date', 'StudentCount'], ascending=[True, False])

        # Save to Excel
        count_df.to_excel(output_file, index=False)
        print(f"✅ Subject-wise student count saved to: {output_file}")

        return count_df
    
    except Exception as e:
        print("❌ Error generating subject-wise student count:", str(e))
        return pd.DataFrame()

def detect_exam_conflicts(df_timetable, df_course_roll_mapping):
    """Detect roll number conflicts in exam schedule"""
    schedule_subject_rolls = defaultdict(dict)
    conflicts = {}

    try:
        for _, row in df_timetable.iterrows():
            date = row['Date']
            for shift in ['Morning', 'Evening']:
                subjects = row.get(shift)
                if isinstance(subjects, str) and subjects.strip().upper() != "NO EXAM":
                    course_codes = [s.strip() for s in subjects.split(';') if s.strip()]
                    for code in course_codes:
                        try:
                            roll_set = set(
                                df_course_roll_mapping[df_course_roll_mapping['course_code'] == code]['rollno'].tolist()
                            )
                            schedule_subject_rolls[(date, shift)][code] = roll_set
                        except Exception as e:
                            logging.error(f"Error fetching roll numbers for course code {code} on {date} ({shift}): {e}")

    except Exception as e:
        logging.exception("Error while processing timetable.")

    try:
        for key, subject_rolls in schedule_subject_rolls.items():
            subjects = list(subject_rolls.keys())
            n = len(subjects)
            for i in range(n):
                for j in range(i + 1, n):
                    code_i = subjects[i]
                    code_j = subjects[j]
                    intersection = subject_rolls[code_i].intersection(subject_rolls[code_j])
                    if intersection:
                        if key not in conflicts:
                            conflicts[key] = []
                        for roll in intersection:
                            try:
                                room_i = df_course_roll_mapping[
                                    (df_course_roll_mapping['course_code'] == code_i) &
                                    (df_course_roll_mapping['rollno'] == roll)
                                ]['room_no'].values

                                room_j = df_course_roll_mapping[
                                    (df_course_roll_mapping['course_code'] == code_j) &
                                    (df_course_roll_mapping['rollno'] == roll)
                                ]['room_no'].values

                                room_i_str = room_i[0] if len(room_i) > 0 else "Unknown"
                                room_j_str = room_j[0] if len(room_j) > 0 else "Unknown"

                                conflicts[key].append(
                                    f"Roll {roll} enrolled in both {code_i} (Room {room_i_str}) and {code_j} (Room {room_j_str})"
                                )
                            except Exception as e:
                                logging.error(f"Error while getting room info for roll {roll}: {e}")
    except Exception as e:
        logging.exception("Error while checking for conflicts.")

    return conflicts

def get_building_priority(building):
    """Get priority for building allocation (B1 gets priority 0, B2 gets priority 1, etc.)"""
    if isinstance(building, str):
        building_upper = building.upper()
        if 'B1' in building_upper:
            return 0
        elif 'B2' in building_upper:
            return 1
        elif 'B3' in building_upper:
            return 2
        else:
            return 999  # Other buildings get lower priority
    return 999

def get_floor_from_room(room_no):
    """Enhanced floor extraction from room number"""
    if pd.isna(room_no):
        return 0
    
    room_str = str(room_no).strip()
    
    # Handle special cases first
    if room_str.startswith('B-'):
        # For rooms like B-001, B-101 (B2 building)
        return 0  # Ground floor
    
    # Extract all digits
    digits = re.findall(r'\d+', room_str)
    if not digits:
        return 0
    
    first_digits = digits[0]
    
    # Determine floor based on room number pattern
    if len(first_digits) == 4:
        # 4-digit rooms (e.g., 6101, 7102)
        return int(first_digits[0])  # First digit is floor
    elif len(first_digits) == 5:
        # 5-digit rooms (e.g., 10502)
        return int(first_digits[:2])  # First two digits are floor
    elif len(first_digits) <= 3:
        # 3-digit or less (e.g., 001, 101)
        return 0  # Ground floor
    
    return 0  # Default case

def allocate_rooms_optimized(subjects_data, rooms_df, buffer, sparse_factor, is_sparse):
    """Optimized room allocation with all requested features"""
    allocation = []
    room_usage = {}
    subject_allocations = defaultdict(list)  # Track where each subject is allocated

    # Prepare rooms data with floor information
    rooms_df = rooms_df.copy()
    rooms_df['Floor'] = rooms_df['Room No.'].apply(get_floor_from_room)
    rooms_df['BuildingPriority'] = rooms_df['Block'].apply(get_building_priority)
    
    # Sort rooms by: building priority (B1 first), then floor, then capacity (descending)
    rooms_df = rooms_df.sort_values(
        by=['BuildingPriority', 'Floor', 'Exam Capacity'],
        ascending=[True, True, False]
    ).reset_index(drop=True)

    # Sort subjects by student count (descending) to prioritize large courses
    subjects_data = sorted(subjects_data, key=lambda x: len(x['students']), reverse=True)

    # Pre-calculate room capacities with buffer
    for _, room in rooms_df.iterrows():
        room_no = room['Room No.']
        capacity = int(room['Exam Capacity'] - buffer)
        allowed_capacity = int(capacity * sparse_factor) if is_sparse else capacity
        room_usage[room_no] = {
            'capacity': allowed_capacity,
            'used': 0,
            'building': room['Block'],
            'floor': room['Floor'],
            'subjects': []
        }

    # Allocation strategy
    for subject_info in subjects_data:
        subject_code = subject_info['subject']
        students = sorted(subject_info['students'])
        students_remaining = len(students)
        
        print(f"\nAllocating {subject_code}: {students_remaining} students")

        # Strategy 1: Try to use rooms where this subject is already partially allocated
        existing_rooms = [
            (room_no, room_info) for room_no, room_info in room_usage.items() 
            if subject_code in room_info['subjects'] and 
               room_info['used'] < room_info['capacity']
        ]
        
        # Sort existing rooms by proximity (same building, then closest floor)
        if subject_allocations.get(subject_code):
            current_building = subject_allocations[subject_code][0]['building']
            current_floor = subject_allocations[subject_code][0]['floor']
            existing_rooms.sort(key=lambda x: (
                0 if x[1]['building'] == current_building else 1,
                abs(x[1]['floor'] - current_floor)
            ))

        # Allocate to existing rooms first
        for room_no, room_info in existing_rooms:
            if students_remaining <= 0:
                break
                
            available = room_info['capacity'] - room_info['used']
            to_allocate = min(available, students_remaining)
            allocated_students = students[:to_allocate]
            
            allocation.append({
                'SubjectCode': subject_code,
                'Room No.': room_no,
                'AssignedCount': to_allocate,
                'AssignedRolls': ';'.join(allocated_students),
                'Building': room_info['building'],
                'Floor': room_info['floor']
            })
            
            room_info['used'] += to_allocate
            students = students[to_allocate:]
            students_remaining -= to_allocate
            print(f"  → Added {to_allocate} to existing room {room_no} (Building: {room_info['building']}, Floor: {room_info['floor']})")

        # Strategy 2: Find new rooms in the same building as existing allocations
        if students_remaining > 0 and subject_allocations.get(subject_code):
            current_building = subject_allocations[subject_code][0]['building']
            building_rooms = [
                (room_no, room_info) for room_no, room_info in room_usage.items()
                if room_info['building'] == current_building and
                   room_info['used'] < room_info['capacity']
            ]
            
            # Sort by floor proximity to existing allocations
            if subject_allocations[subject_code]:
                avg_floor = sum(a['floor'] for a in subject_allocations[subject_code]) / len(subject_allocations[subject_code])
                building_rooms.sort(key=lambda x: abs(x[1]['floor'] - avg_floor))
            
            for room_no, room_info in building_rooms:
                if students_remaining <= 0:
                    break
                    
                available = room_info['capacity'] - room_info['used']
                to_allocate = min(available, students_remaining)
                allocated_students = students[:to_allocate]
                
                allocation.append({
                    'SubjectCode': subject_code,
                    'Room No.': room_no,
                    'AssignedCount': to_allocate,
                    'AssignedRolls': ';'.join(allocated_students),
                    'Building': room_info['building'],
                    'Floor': room_info['floor']
                })
                
                room_info['used'] += to_allocate
                room_info['subjects'].append(subject_code)
                students = students[to_allocate:]
                students_remaining -= to_allocate
                subject_allocations[subject_code].append({
                    'building': room_info['building'],
                    'floor': room_info['floor']
                })
                print(f"  → Added {to_allocate} to same-building room {room_no} (Floor: {room_info['floor']})")

        # Strategy 3: Find any available rooms (fallback)
        for room_no, room_info in room_usage.items():
            if students_remaining <= 0:
                break
                
            if room_info['used'] < room_info['capacity']:
                available = room_info['capacity'] - room_info['used']
                to_allocate = min(available, students_remaining)
                allocated_students = students[:to_allocate]
                
                allocation.append({
                    'SubjectCode': subject_code,
                    'Room No.': room_no,
                    'AssignedCount': to_allocate,
                    'AssignedRolls': ';'.join(allocated_students),
                    'Building': room_info['building'],
                    'Floor': room_info['floor']
                })
                
                room_info['used'] += to_allocate
                room_info['subjects'].append(subject_code)
                students = students[to_allocate:]
                students_remaining -= to_allocate
                subject_allocations[subject_code].append({
                    'building': room_info['building'],
                    'floor': room_info['floor']
                })
                print(f"  → Added {to_allocate} to new room {room_no} (Building: {room_info['building']}, Floor: {room_info['floor']})")

        if students_remaining > 0:
            print(f"⚠️ WARNING: {students_remaining} students from {subject_code} could not be allocated")

    # Room utilization summary
    print("\n📊 Optimized Room Utilization Summary:")
    for room_no, info in room_usage.items():
        if info['used'] > 0:
            utilization = (info['used'] / info['capacity']) * 100 if info['capacity'] > 0 else 0
            subjects_str = ', '.join(info['subjects'])
            print(f"  {room_no} ({info['building']}, Floor {info['floor']}): "
                  f"{info['used']}/{info['capacity']} ({utilization:.1f}%) - Subjects: {subjects_str}")

    return allocation

def create_excel_with_ta_invigilator_footer(df_output, heading_text, filepath, subject_list=None):
    """Create Excel file with TA and Invigilator rows at the footer, supports multiple subjects"""
    try:
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Attendance Sheet')
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#D7E4BC',
                'border': 1
            })
            
            subject_header_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#E6E6E6',
                'border': 1
            })
            
            footer_format = workbook.add_format({
                'bg_color': '#F2F2F2',
                'border': 1,
                'align': 'center',
                'bold': True
            })
            
            student_format = workbook.add_format({
                'border': 1
            })
            
            # Write main header
            worksheet.merge_range('A1:C1', heading_text, header_format)
            
            # Write column headers
            worksheet.write('A2', 'Roll No.', header_format)
            worksheet.write('B2', 'Student Name', header_format)
            worksheet.write('C2', 'Signature', header_format)
            
            # Start row for student data
            current_row = 2
            
            # If multiple subjects, add subject headers
            if subject_list and len(subject_list) > 1:
                for subject in subject_list:
                    current_row += 1
                    worksheet.merge_range(
                        current_row, 0, current_row, 2, 
                        f"Subject: {subject}", 
                        subject_header_format
                    )
                    
                    # Get students for this subject
                    subject_students = df_output[df_output['Subject'] == subject]
                    
                    # Write student data
                    for _, row in subject_students.iterrows():
                        current_row += 1
                        worksheet.write(current_row, 0, str(row['Roll']), student_format)
                        worksheet.write(current_row, 1, str(row.get('Student Name', '')), student_format)
                        worksheet.write(current_row, 2, '', student_format)
            else:
                # Single subject case
                for _, row in df_output.iterrows():
                    current_row += 1
                    worksheet.write(current_row, 0, str(row['Roll']), student_format)
                    worksheet.write(current_row, 1, str(row.get('Student Name', '')), student_format)
                    worksheet.write(current_row, 2, '', student_format)
            
            # Add some spacing before footer
            current_row += 2
            
            # Add footer header
            worksheet.merge_range(current_row, 0, current_row, 2, 'EXAMINATION STAFF', header_format)
            current_row += 1
            
            # Add TA rows (as footer)
            for i in range(1, 6):
                current_row += 1
                worksheet.write(current_row, 0, f'TA{i}', footer_format)
                worksheet.write(current_row, 1, '', footer_format)
                worksheet.write(current_row, 2, '', footer_format)
            
            # Add Invigilator rows (as footer)
            for i in range(1, 6):
                current_row += 1
                worksheet.write(current_row, 0, f'Invigilator{i}', footer_format)
                worksheet.write(current_row, 1, '', footer_format)
                worksheet.write(current_row, 2, '', footer_format)
            
            # Set column widths
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:B', 25)
            worksheet.set_column('C:C', 20)
            
    except Exception as e:
        print(f"❌ Error creating Excel file {filepath}: {e}")


def perform_room_allocation(
    df_timetable,
    df_course_roll_mapping,
    df_roll_name_mapping,
    df_room_capacity,
    count_df,
    buffer,
    sparse_factor,
    is_sparse,
    output_base_folder="output_folder_roomwise"
):
    """Perform room allocation based on user preferences"""
    
    # Prepare Room Capacity with improved sorting
    df_room_capacity['Room No.'] = df_room_capacity['Room No.'].astype(str).str.strip()
    df_room_capacity['Floor'] = df_room_capacity['Room No.'].apply(get_floor_from_room)
    df_room_capacity['BuildingPriority'] = df_room_capacity['Block'].apply(get_building_priority)
    df_room_capacity['Exam Capacity'] = df_room_capacity['Exam Capacity'].astype(int)
    
    # Sort rooms by building priority (B1 first), then floor, then capacity (descending)
    df_room_capacity_sorted = df_room_capacity.sort_values(
        by=['BuildingPriority', 'Floor', 'Exam Capacity'],
        ascending=[True, True, False]
    ).reset_index(drop=True)
    
    final_allocation = []

    for _, row in df_timetable.iterrows():
        date = row['Date']
        for shift in ['Morning', 'Evening']:
            if pd.notnull(row[shift]) and row[shift].strip().upper() != "NO EXAM":
                subjects = [s.strip() for s in row[shift].split(';') if s.strip()]
                
                # Sort subjects by student count (from count_df)
                subjects_with_counts = []
                for subject_code in subjects:
                    count = count_df[(count_df['Date'] == date) & 
                                   (count_df['SubjectCode'] == subject_code)]['StudentCount'].values
                    if len(count) > 0:
                        subjects_with_counts.append((subject_code, count[0]))
                    else:
                        subjects_with_counts.append((subject_code, 0))
                
                # Sort by student count (descending)
                subjects_with_counts.sort(key=lambda x: x[1], reverse=True)
                subjects_sorted = [x[0] for x in subjects_with_counts]
                
                # Prepare subject data
                subjects_data = []
                for subject_code in subjects_sorted:
                    student_rolls = df_course_roll_mapping[
                        df_course_roll_mapping['course_code'] == subject_code
                    ]['rollno'].unique()
                    subjects_data.append({
                        'subject': subject_code,
                        'students': list(student_rolls)
                    })

                # Use the optimized allocation function
                allocation = allocate_rooms_optimized(
                    subjects_data, 
                    df_room_capacity_sorted, 
                    buffer, 
                    sparse_factor, 
                    is_sparse
                )

                for alloc in allocation:
                    final_allocation.append({
                        'Date': date,
                        'Shift': shift,
                        'SubjectCode': alloc['SubjectCode'],
                        'Room No.': alloc['Room No.'],
                        'AssignedCount': alloc['AssignedCount'],
                        'AssignedRolls': alloc['AssignedRolls'],
                        'Building': alloc['Building'],
                        'Floor': alloc['Floor']
                    })

    # Save Final Allocation
    df_final_allocation = pd.DataFrame(final_allocation)
    df_final_allocation.to_excel("final_room_allocation.xlsx", index=False)
    print(f"✅ Final allocation saved to: final_room_allocation.xlsx")

    # Export Room-wise Sheets - CORRECTED FOLDER CREATION
    try:
        # Create base output folder if it doesn't exist
        os.makedirs(output_base_folder, exist_ok=True)
        print(f"📁 Created base output folder: {output_base_folder}")
        
        # Group by Date, Shift, and Room No.
        group_cols = ['Date', 'Shift', 'Room No.']
        grouped = df_final_allocation.groupby(group_cols)

        for group_keys, df_group in grouped:
            date, shift, room = group_keys
            try:
                # Convert date to proper string format
                date_str = pd.to_datetime(date).strftime("%Y-%m-%d")
                
                # Create date folder path
                date_folder = os.path.join(output_base_folder, date_str)
                os.makedirs(date_folder, exist_ok=True)
                
                # Create shift folder path
                shift_folder = os.path.join(date_folder, shift)
                os.makedirs(shift_folder, exist_ok=True)
                
                # Get all assigned roll numbers with their subjects
                all_rolls = []
                subjects_in_room = set()
                
                for _, row in df_group.iterrows():
                    rolls = str(row['AssignedRolls']).split(';')
                    for roll in rolls:
                        if roll.strip():
                            all_rolls.append({
                                'Roll': roll.strip(),
                                'Subject': row['SubjectCode']
                            })
                            subjects_in_room.add(row['SubjectCode'])
                
                # Create output dataframe
                df_output = pd.DataFrame(all_rolls)
                
                # Handle case where roll-name mapping might have different column names
                roll_name_col = 'Name' if 'Name' in df_roll_name_mapping.columns else 'Student Name'
                roll_col = 'rollno' if 'rollno' in df_roll_name_mapping.columns else 'Roll'
                
                df_output['Student Name'] = df_output['Roll'].map(
                    df_roll_name_mapping.set_index(roll_col)[roll_name_col]
                ).fillna('')
                
                # Sort by subject then by roll number
                df_output = df_output.sort_values(['Subject', 'Roll'])
                
                # Create heading
                subjects_list = sorted(list(subjects_in_room))
                heading_text = f"Room: {room} | Date: {date_str} | Session: {shift} | Subjects: {', '.join(subjects_list)}"
                filename = f"Room_{room}.xlsx"
                filepath = os.path.join(shift_folder, filename)

                # Create Excel with all subjects in one sheet
                create_excel_with_ta_invigilator_footer(
                    df_output, 
                    heading_text, 
                    filepath, 
                    subject_list=subjects_list if len(subjects_list) > 1 else None
                )
                
                print(f"✅ Saved Room sheet → {filepath}")
                
            except Exception as e:
                print(f"❌ Error creating room sheet for {date} {shift} {room}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"❌ Error creating output folder structure: {str(e)}")
        return

    print("\n🎉 All room-wise sheets generated successfully!")


def main():
    """Main function"""
    try:
        print("🎓 EXAM SEATING ARRANGEMENT SYSTEM (OPTIMIZED)")
        print("=" * 50)
        print("✨ Optimization Features:")
        print("   • Largest courses allocated first")
        print("   • Subjects kept within same building")
        print("   • Rooms on same/similar floors prioritized")
        print("   • B1 building gets highest priority")
        print("   • Buffer and Sparse/Dense options implemented")
        print("   • Single sheet per room for multiple subjects")
        
        # Get user inputs
        buffer, sparse_factor, is_sparse = get_user_inputs()
        
        # Load Input Excel
        print("\n📖 Loading input data...")
        df_timetable = pd.read_excel("input_data_tt.xlsx", sheet_name="in_timetable")
        df_course_roll_mapping = pd.read_excel("input_data_tt.xlsx", sheet_name="in_course_roll_mapping")
        df_roll_name_mapping = pd.read_excel("input_data_tt.xlsx", sheet_name="in_roll_name_mapping")
        df_room_capacity = pd.read_excel("input_data_tt.xlsx", sheet_name="in_room_capacity")

        # Clean Timetable Columns
        df_timetable['Morning'] = df_timetable['Morning'].apply(clean_timetable)
        df_timetable['Evening'] = df_timetable['Evening'].apply(clean_timetable)

        # Conflict Detection
        print("\n🔍 Detecting exam conflicts...")
        conflicts = detect_exam_conflicts(df_timetable, df_course_roll_mapping)

        if conflicts:
            print("⚠️ CONFLICTS DETECTED:")
            for (date, shift), messages in conflicts.items():
                logging.warning(f"[Exam Date: {date}] Conflict during {shift} shift:")
                print(f"📅 Conflict on {date} during {shift} shift:")
                for msg in messages:
                    logging.warning(f"[Exam Date: {date}] {msg}")
                    print(f"    ❌ {msg}")
        else:
            logging.info("No roll number conflicts found.")
            print("✅ No roll number conflicts found.")

        # Subject-wise Student Count
        print("\n📊 Generating subject-wise student count...")
        count_df = generate_subjectwise_student_count(df_timetable, df_course_roll_mapping)
        
        if count_df.empty:
            print("❌ Failed to generate subject-wise student count. Check logs for details.")
            return
        
        # Room Allocation
        print("\n🏫 Performing optimized room allocation...")
        perform_room_allocation(
            df_timetable,
            df_course_roll_mapping,
            df_roll_name_mapping,
            df_room_capacity,
            count_df,
            buffer,
            sparse_factor,
            is_sparse
        )
        
        print("\n🎉 PROCESS COMPLETED SUCCESSFULLY!")
        print("📁 Check the 'output_folder_roomwise' directory for room-wise Excel sheets.")
        print("📋 Check 'final_room_allocation.xlsx' for complete allocation summary.")
        print("\n🆕 Optimization Features Applied:")
        print("   ✅ Largest courses allocated first to minimize room usage")
        print("   ✅ Subjects kept within same building when possible")
        print("   ✅ Rooms on same/similar floors prioritized")
        print("   ✅ B1 building gets highest priority")
        print("   ✅ Buffer and Sparse/Dense options implemented")
        print("   ✅ Single sheet per room for multiple subjects")

    except Exception as e:
        logging.exception("❌ Error in main function.")
        print(f"❌ An error occurred: {str(e)}")
        print("📋 Check 'exam_conflicts.log' for detailed error information.")

if __name__ == "__main__":
    main()