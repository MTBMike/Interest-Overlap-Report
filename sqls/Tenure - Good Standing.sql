with tenures as
(
SELECT 
  /* Title Attributes */
  t.tenure_number_id, 
  TRIM (tetc.description) tenure_type_description,
  TRIM (tstc.description) tenure_sub_type_description,
  TRIM (titc.description) title_type_description,
  T.issue_date issue_date, 
  T.good_to_date good_to_date,
  T.area_in_hectares, 
  T.protected_ind,
  /* Client Attributes */
  LISTAGG(ctx.client_number_id || ': ' || TRIM(povw.NAME) || ' ' || ctx.tenure_percentage || '%' , ', ') WITHIN GROUP (ORDER BY ctx.client_number_id) AS TENURE_OWNERS,
  (
   SELECT   
    COUNT (ctx1.client_number_id)
   FROM 
    mta_acquired_tenure_poly atp1 INNER JOIN mta_client_tenure_xref ctx1
    ON atp1.tenure_number_id = ctx1.tenure_number_id
   WHERE 
    t.tenure_number_id = atp1.tenure_number_id
   GROUP BY 
    atp1.tenure_number_id
  ) number_of_owners,
  /* Other Dataset Attributes */
  T.termination_date,
  TRIM (ttc.description) termination_type_description
FROM 
  mta.mta_tenure t
  INNER JOIN mta.mta_tenure_type_code tetc
  ON T.mta_tenure_type_code = tetc.mta_tenure_type_code
  INNER JOIN mta.mta_tenure_sub_type_code tstc
  ON T.mta_tenure_sub_type_code = tstc.mta_tenure_sub_type_code
  INNER JOIN mta.mta_title_type_code titc
  ON T.mta_title_type_code = titc.mta_title_type_code
  LEFT OUTER JOIN mta.mta_termination_type_code ttc
  ON T.mta_termination_type_code = ttc.mta_termination_type_code
  INNER JOIN mta.mta_client_tenure_xref ctx
  ON t.tenure_number_id = ctx.tenure_number_id
  INNER JOIN mta.mta_person_organization_vw povw
  ON ctx.client_number_id = povw.client_number_id
where
  povw.client_number_id IS NOT NULL
group by
  t.tenure_number_id, 
  TRIM (t.claim_name),
  TRIM (tetc.description),
  TRIM (tstc.description),
  TRIM (titc.description),
  T.issue_date, 
  T.good_to_date,
  T.area_in_hectares, 
  T.protected_ind, 
  T.termination_date,
  TRIM (ttc.description)
)

select 
  atp.objectid,
  t.tenure_number_id, 
  t.tenure_type_description,
  t.tenure_sub_type_description,
  t.title_type_description,
  t.issue_date issue_date, 
  t.good_to_date good_to_date,
  t.area_in_hectares, 
  t.protected_ind,
  t.TENURE_OWNERS,
  t.number_of_owners,
  t.termination_date,
  t.termination_type_description,
  atp.geometry,
  atp.geometry_area,
  atp.geometry_len
from
  tenures t
  inner join mta_acquired_tenure_poly atp 
  on t.tenure_number_id = atp.tenure_number_id
where
  update_query
order by
  t.tenure_number_id


