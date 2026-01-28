from django.utils import timezone
from django.db import models
from datetime import datetime
from calendar import monthrange

class Workshop(models.Model): 
    name = models.CharField(max_length=150)
    name_kz = models.CharField(max_length=150, null=True, blank=True)
    code = models.IntegerField(unique=True, null=True, blank=True)
    ord_s = models.IntegerField(null=True, blank=True)
    
    def __str__(self):
        return self.name

class Resource(models.Model): 
    name = models.CharField(max_length=150, null=True, blank=True)
    name_kz = models.CharField(max_length=150, null=True, blank=True)
    unit = models.CharField(max_length=20)
    priority = models.IntegerField(default=0, null=True, blank=True)
    
    def __str__(self):
        return f"{self.name} ({self.unit})"

class Plan(models.Model):
    workshop = models.ForeignKey(
        Workshop,
        on_delete=models.CASCADE,
        related_name='plans'
    )
    resource = models.ForeignKey(
        Resource,
        on_delete=models.CASCADE,
        related_name='plans'
    )
    year = models.IntegerField()
    month = models.IntegerField()  
    plan_value = models.FloatField(help_text="План")
    
    class Meta:
        unique_together = ['workshop', 'resource', 'year', 'month']
        ordering = ['-year', '-month']
    
    def __str__(self):
        return f"{self.workshop.name} - {self.resource.name} ({self.year}-{self.month:02d}): {self.plan_value}"
    
    def get_daily_plan(self):
        days_in_month = monthrange(self.year, self.month)[1]
        return self.plan_value / days_in_month
    
    def get_plan_for_period(self, start_date, end_date):
        days_in_month = monthrange(self.year, self.month)[1]
        daily_plan = self.plan_value / days_in_month

        month_start = datetime(self.year, self.month, 1).date()
        month_end = datetime(self.year, self.month, days_in_month).date()
        
        period_start = max(start_date, month_start)
        period_end = min(end_date, month_end)
        
        if period_start > period_end:
            return 0
        
        days_count = (period_end - period_start).days + 1
        return daily_plan * days_count

class DailyFact(models.Model):
    workshop = models.ForeignKey(
        Workshop,
        on_delete=models.CASCADE,
        related_name='daily_facts'
    )
    resource = models.ForeignKey(
        Resource,
        on_delete=models.CASCADE,
        related_name='daily_facts'
    )
    date = models.DateField(default=timezone.now)
    fact_value = models.FloatField(help_text="Фактическое значение за день")
    
    class Meta:
        unique_together = ['workshop', 'resource', 'date']
        ordering = ['-date']
    
    def __str__(self):
        return f"{self.workshop.name} - {self.resource.name} ({self.date}): {self.fact_value}"

class StatisticsManager:
    
    @staticmethod
    def get_monthly_statistics(workshop, resource, year, month):
        from datetime import date
        
        try:
            monthly_plan = Plan.objects.get(
                workshop=workshop,
                resource=resource,
                year=year,
                month=month
            )
            plan_month = monthly_plan.plan_value
            plan_daily = monthly_plan.get_daily_plan()
        except Plan.DoesNotExist:
            plan_month = 0
            plan_daily = 0
        
        days_in_month = monthrange(year, month)[1]
        start_date = date(year, month, 1)
        end_date = date(year, month, days_in_month)
        
        daily_facts = DailyFact.objects.filter(
            workshop=workshop,
            resource=resource,
            date__range=[start_date, end_date]
        )
        
        fact_month = sum(df.fact_value for df in daily_facts)
        
        today = timezone.now().date()
        if today.year == year and today.month == month:
            current_day = today.day
        else:
            current_day = days_in_month
        
        plan_to_date = plan_daily * current_day
        
        return {
            'plan_month': plan_month,
            'plan_daily': plan_daily,
            'plan_to_date': plan_to_date,
            'fact_month': fact_month,
            'days_in_month': days_in_month,
            'current_day': current_day,
            'deviation': fact_month - plan_to_date,
            'completion_percent': (fact_month / plan_to_date * 100) if plan_to_date > 0 else 0
        }
    @staticmethod
    def get_yearly_statistics(workshop, resource, year, current_month):
        """
        ГОД = закрытые месяцы + текущий месяц ДО СЕГОДНЯ
        """

        from django.utils import timezone

        today = timezone.now().date()

        total_plan = 0
        total_fact = 0

        for m in range(1, current_month):
            stats = StatisticsManager.get_monthly_statistics(
                workshop, resource, year, m
            )
            total_plan += stats['plan_month']
            total_fact += stats['fact_month']

        current_stats = StatisticsManager.get_monthly_statistics(
            workshop, resource, year, current_month
        )

        total_plan += current_stats['plan_to_date']
        total_fact += current_stats['fact_month']

        return {
            'year_plan_to_date': total_plan,
            'year_fact_to_date': total_fact,
            'year_deviation': total_fact - total_plan,
            'year_percent': (total_fact / total_plan * 100) if total_plan > 0 else 0
        }
        
    @staticmethod
    def get_daily_statistics(workshop, resource, target_date):
        """Получает статистику за конкретный день"""
        # План на день
        try:
            monthly_plan = Plan.objects.get(
                workshop=workshop,
                resource=resource,
                year=target_date.year,
                month=target_date.month
            )
            plan_daily = monthly_plan.get_daily_plan()
        except Plan.DoesNotExist:
            plan_daily = 0
        
        # Факт за день
        try:
            daily_fact = DailyFact.objects.get(
                workshop=workshop,
                resource=resource,
                date=target_date
            )
            fact_daily = daily_fact.fact_value
        except DailyFact.DoesNotExist:
            fact_daily = 0
        
        return {
            'date': target_date,
            'plan_daily': plan_daily,
            'fact_daily': fact_daily,
            'deviation': fact_daily - plan_daily,
            'completion_percent': (fact_daily / plan_daily * 100) if plan_daily > 0 else 0
        }