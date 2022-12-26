package ru.m9ist;

import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import jxl.write.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Objects;
import java.util.stream.Collectors;

public class CalendarGenerator {
    private static final int YEAR = 2023;

    public static void main(final String[] args) throws IOException, BiffException, WriteException {
        final InputStream orStream = CalendarGenerator.class.getClassLoader().getResourceAsStream("schedule2020.xls");
        final File newFile = new File("schedule" + YEAR + "_generated.xls");
        final WritableWorkbook workbook = Workbook.createWorkbook(newFile, Workbook.getWorkbook(Objects.requireNonNull(orStream)));
        final WritableSheet sheet = workbook.getSheet(0);
        int i = 26;
        final GregorianCalendar gregorianCalendar = new GregorianCalendar(YEAR, Calendar.JANUARY, 1);
        getNextWorkDay(gregorianCalendar);
        gregorianCalendar.add(Calendar.DAY_OF_YEAR, -1);
        final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd.MM.yyyy");
        final ArrayList<Integer> days = new ArrayList<>();
        while (true) {
            final EndWeek nextWorkDay = getNextWorkDay(gregorianCalendar);
            i++;
            if (nextWorkDay != EndWeek.NO) {
                endWeek(sheet, days, gregorianCalendar, i, nextWorkDay != EndWeek.NEED_WEEK_END);
                i++;
            }
            if (nextWorkDay == EndWeek.NEED_YEAR_END) break;

            days.add(i);
            final CellFormat cellFormat = sheet.getCell("E4").getCellFormat();
            sheet.addCell(new Label(0, i, simpleDateFormat.format(gregorianCalendar.getTime())));
            sheet.addCell(new Formula(3, i, String.format("B%d+$A$3+F%1$d", i + 1), cellFormat));
            sheet.addCell(new Formula(4, i, String.format("C%d-B%1$d-F%1$d", i + 1), cellFormat));
        }
        orStream.close();
        workbook.write();
        workbook.close();
    }

    private enum EndWeek {
        NO,
        NEED_WEEK_END,
        NEED_MONTH_END,
        NEED_YEAR_END,
    }

    private static EndWeek getNextWorkDay(final GregorianCalendar gregorianCalendar) {
        final int prevMonth = gregorianCalendar.get(Calendar.MONTH);
        boolean wasWeekEnd = false;
        while (true) {
            gregorianCalendar.add(Calendar.DAY_OF_YEAR, 1);
            if (gregorianCalendar.get(Calendar.YEAR) != YEAR) {
                return EndWeek.NEED_YEAR_END;
            }
            final int weekDay = gregorianCalendar.get(Calendar.DAY_OF_WEEK);
            if (weekDay == Calendar.SATURDAY || weekDay == Calendar.SUNDAY) {
                wasWeekEnd = true;
                continue;
            }
            if (prevMonth == gregorianCalendar.get(Calendar.MONTH)) {
                return wasWeekEnd ? EndWeek.NEED_WEEK_END : EndWeek.NO;
            } else {
                return EndWeek.NEED_MONTH_END;
            }
        }
    }

    private static final ArrayList<Integer> weeksEnd = new ArrayList<>();

    private static void endWeek(final WritableSheet sheet, final ArrayList<Integer> days, final GregorianCalendar gregorianCalendar, final int i, final boolean isMonthEnd) throws WriteException {
        weeksEnd.add(i);
        if (days.isEmpty()) throw new IllegalStateException();
        final String daysSum = days.stream().map(day -> "E" + (day + 1)).collect(Collectors.joining("+"));
        sheet.addCell(new Formula(4, i, daysSum, sheet.getCell(4, 6).getCellFormat()));
        if (isMonthEnd) {
            if (weeksEnd.isEmpty()) throw new IllegalStateException();
            final String weeksSum = weeksEnd.stream().map(week -> "E" + (week + 1)).collect(Collectors.joining("+"));
            sheet.addCell(new Formula(5, i, weeksSum, sheet.getCell(5, 23).getCellFormat()));
            final GregorianCalendar gregorianCalendarCopy = new GregorianCalendar();
            gregorianCalendarCopy.setTime(gregorianCalendar.getTime());
            gregorianCalendarCopy.add(Calendar.MONTH, -1);
            sheet.addCell(new Label(6, i, new SimpleDateFormat("MM.yyyy").format(gregorianCalendarCopy.getTime())));
            weeksEnd.clear();
        }
        days.clear();
    }
}
