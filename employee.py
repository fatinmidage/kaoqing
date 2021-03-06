class Employee:
    def __init__(self, name, six_workday_mode=False):
      self.__name = name
      self.__yingchu_days = 0
      self.__actual_days = 0
      self.__personal_leave_days = 0
      self.__holidays = 0
      # self.__delay_count = 0
      # self.__buka_count = 0
      self.__quanqing = True
      self.__six_day_mode = six_workday_mode
    
    def get_name(self):
      return self.__name

    def set_six_workdays_mode(self):
      self.__six_day_mode = True
    
    def get_workdays_mode(self):
      return self.__six_day_mode
    
    def get_yingchu_days(self):
      return self.__yingchu_days

    def set_yingchu_days(self,days):
      self.__yingchu_days = days

    def add_actual_workdays(self,):
      self.__actual_days +=1

    def get_actual_workdays(self):
      return self.__actual_days

    def get_quanqing(self):
      return self.__quanqing
    
    def set_quanqing(self,quanqing):
      self.__quanqing = quanqing

    def get_holidays(self):
      return self.__holidays
    
    def add_holidays(self,holidays):
      self.__holidays += holidays