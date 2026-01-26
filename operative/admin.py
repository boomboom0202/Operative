from django.contrib import admin
from django.utils.html import format_html
from django.db.models import Sum
from django.urls import path
from django.shortcuts import render
from datetime import date, timezone
from operative.models import DailyFact, MonthlyPlan, Resource, StatisticsManager, Workshop

@admin.register(Workshop)
class WorkshopAdmin(admin.ModelAdmin):
    list_display = ['name']
    search_fields = ['name']


@admin.register(Resource)
class ResourceAdmin(admin.ModelAdmin):
    list_display = ['name', 'unit']
    search_fields = ['name']
    list_filter = ['unit']


@admin.register(MonthlyPlan)
class MonthlyPlanAdmin(admin.ModelAdmin):
    list_display = ['workshop', 'resource', 'year', 'month', 'plan_value', 'daily_plan_display']
    list_filter = ['workshop', 'resource', 'year', 'month']
    search_fields = ['workshop__name', 'resource__name']
    ordering = ['-year', '-month']
    
    def daily_plan_display(self, obj):
        return f"{obj.get_daily_plan():.2f}"
    daily_plan_display.short_description = 'Дневной план'
    
    fieldsets = (
        ('Основная информация', {
            'fields': ('workshop', 'resource')
        }),
        ('Период', {
            'fields': ('year', 'month')
        }),
        ('План', {
            'fields': ('plan_value',)
        }),
    )


@admin.register(DailyFact)
class DailyFactAdmin(admin.ModelAdmin):
    list_display = ['workshop', 'resource', 'date', 'fact_value', 'plan_comparison']
    list_filter = ['workshop', 'resource', 'date']
    search_fields = ['workshop__name', 'resource__name']
    date_hierarchy = 'date'
    ordering = ['-date']
    
    def plan_comparison(self, obj):
        """Показывает сравнение с планом"""
        try:
            monthly_plan = MonthlyPlan.objects.get(
                workshop=obj.workshop,
                resource=obj.resource,
                year=obj.date.year,
                month=obj.date.month
            )
            daily_plan = monthly_plan.get_daily_plan()
            deviation = obj.fact_value - daily_plan
            
            if deviation >= 0:
                color = 'green'
                sign = '+'
            else:
                color = 'red'
                sign = ''
            
            return format_html(
                '<span style="color: {};">{}{:.2f} ({:.1f}%)</span>',
                color,
                sign,
                deviation,
                (obj.fact_value / daily_plan * 100) if daily_plan > 0 else 0
            )
        except MonthlyPlan.DoesNotExist:
            return format_html('<span style="color: gray;">Нет плана</span>')
    
    plan_comparison.short_description = 'Отклонение от плана'
    
    fieldsets = (
        ('Основная информация', {
            'fields': ('workshop', 'resource')
        }),
        ('Дата и значение', {
            'fields': ('date', 'fact_value')
        }),
    )
    
    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path('statistics-dashboard/', self.admin_site.admin_view(self.statistics_dashboard), name='statistics-dashboard'),
        ]
        return custom_urls + urls
    
    def statistics_dashboard(self, request):
        today = timezone.now().date()
        year = int(request.GET.get('year', today.year))
        month = int(request.GET.get('month', today.month))
        
        workshops = Workshop.objects.all()
        resources = Resource.objects.all()
        
        statistics = []
        for workshop in workshops:
            for resource in resources:
                stats = StatisticsManager.get_monthly_statistics(
                    workshop, resource, year, month
                )
                if stats['plan_month'] > 0:  # Показываем только там, где есть план
                    statistics.append({
                        'workshop': workshop,
                        'resource': resource,
                        'year': year,
                        'month': month,
                        **stats
                    })
        
        context = {
            **self.admin_site.each_context(request),
            'title': 'Статистика плана и факта',
            'statistics': statistics,
            'year': year,
            'month': month,
            'months': range(1, 13),
            'years': range(today.year - 2, today.year + 2),
        }
        
        return render(request, 'admin/statistics_dashboard.html', context)