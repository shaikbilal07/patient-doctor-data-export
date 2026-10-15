SELECT
    p.patient_id,
    p.patient_name,
    p.age            AS patient_age,
    p.gender         AS patient_gender,
    p.city           AS patient_city,
    p.admitted_on,
    p.doctor_id,
    d.doctor_name,
    d.specialization AS doctor_specialization,
    d.phone          AS doctor_phone
FROM patient_details AS p
LEFT JOIN doctor AS d ON d.doctor_id = p.doctor_id
ORDER BY p.patient_id
