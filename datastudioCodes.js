//days till end of month
DATE_DIFF(
	DATETIME_SUB(
		DATETIME_ADD(
			date(year(today()),month(today()),1),
		INTERVAL 1 MONTH
		), 
		INTERVAL 1 DAY 
	),
	date(year(today()),month(today()),1)
)

//days till yesterday
date_diff(
	datetime_sub(
		today(),
		INTERVAL 1 DAY
	),
	Date(year(today()),month(today()),1)
)

// days left till the end of the month
DATE_DIFF(
  DATETIME_SUB(
    datetime_add(
      DATE(year(today()),month(today()),1), 
      INTERVAL 1 MONTH
    ),
    INTERVAL 1 DAY
  ),
  datetime_SUB(
    today(),
    INTERVAL 1 DAY
  )
)


// average spend per day
sum(Cost) 
/
max(
  (date_diff(
    datetime_sub(
      today(),
      INTERVAL 1 DAY
    ),
    Date(
      year(today()),
      month(today()),
      1
    )
  ) +1 )
)

// target spend per day
(max(Budget)-sum(Cost)) 
/
Max(
	DATE_DIFF(
		DATETIME_SUB(
			datetime_add(
				DATE(year(today()),month(today()),1),
				INTERVAL 1 MONTH
			),INTERVAL 1 DAY
		),
		datetime_SUB(
			today(),
			INTERVAL 1 DAY
		)
	)
)

// budget forecast
(
	sum(Cost) 
	/ 
	(date_diff(
		datetime_sub(
			today(),
			INTERVAL 1 DAY
		),
		Date(
			year(today()),
			month(today()),
			1
		)
	)+1)
) * 
DATE_DIFF(
	DATETIME_SUB(
		datetime_add(
			DATE(year(today()),month(today()),1), 
			INTERVAL 1 MONTH
		),
		INTERVAL 1 DAY),
	datetime_SUB(
		DATE(year(today()),month(today()),1),
		INTERVAL 1 DAY
	)
)
