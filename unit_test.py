from Defclass import *
import unittest

class LimesurveyExcelDatasourceTest(unittest.TestCase):
    def test_limesurvey_simple_datasource(self):
        a=LimesurveyExcelDatasource('LS_1BESOIN','1BESOIN.xlsx').get_var('submitdate')
        self.assertEqual(a.var_name,'submitdate')
        self.assertEqual(a.value.value,'1980-01-01 00:00:00')
        self.assertEqual(a.value.value_type,ValueType.STRING)
        
    def test_ls_simple_var_not_found(self):
        try :
            LimesurveyMultiExcelDatasource('HRMI01','LS_BESOIN','1BESOIN.xlsx').get_var('CODE')
        except ValueError :
            pass
        else :
            self.fail("Il devrait y'avoir une erreur de type ValueError")
        

class LimesurveyExcelMultiDatasourceTestIncrement(unittest.TestCase):
    def test_limesurvey_multi_datasource(self):
        
        a = LimesurveyMultiExcelDatasource('HRMI01','LS_2ADHESION','2ADHESION.xlsx')
        
        b = a.get_var('CODE_FORMATION')
        self.assertEqual(b.value.value,'HRMI0101')
        
        c = a.next_item().get_var('CODE_FORMATION')
        self.assertEqual(c.value.value,'HRMI0102')
        
        d = a.next_item().get_var('CODE_FORMATION')
        self.assertEqual(d.value.value,'HRMI0103')
        
        e = a.next_item()
        self.assertEqual(e,None)


class LimesurveyExcelMultiDatasourceTestClass(unittest.TestCase):
    def test_limesurvey_multi_datasource(self):
        a = LimesurveyMultiExcelDatasource('HRMI01','LS_2ADHESION','2ADHESION.xlsx').get_var('CODE_FORMATION')
        self.assertEqual(a.var_name,'CODE_FORMATION')
        self.assertEqual(a.value.value,'HRMI0101')
        self.assertEqual(a.value.value_type,ValueType.STRING)
        
    def test_ls_multi_var_not_found(self):
        try :
            LimesurveyMultiExcelDatasource('HRMI01','LS_2ADHESION','2ADHESION.xlsx').get_var('CODE')
        except ValueError :
            pass
        else :
            self.fail("Il devrait y'avoir une erreur de type ValueError")
        

class InfoclientDatasourceTest(unittest.TestCase):
    def test_infoclient_datasource(self):
        a=InfoClientDatasource('Fiche_Client','HRMI01.xlsx').get_var('Acpte')
        self.assertEqual(a.var_name,'Acpte')
        self.assertEqual(a.value.value,0.5)
        self.assertEqual(a.value.value_type,ValueType.NUMERIC)
        
    def test_infoclient_named_zone_not_found(self):
        try :
            InfoClientDatasource('Fiche_client', 'HRMI01.xlsx').get_var('titi')
        except ValueError :
            pass
        else :
            self.fail("Il devrait y'avoir une erreur de type ValueError")
        

class DatasourceManagerTest(unittest.TestCase):
    def test_datasource_manager(self):
        a = DatasourceManager('test.xlsx','Source','HRMI01').get_var('LS_1BESOIN','submitdate')
        self.assertEqual(a.value.value,'1980-01-01 00:00:00')
        self.assertEqual(a.value.value_type,ValueType.STRING)
        self.assertEqual(a.var_name,'submitdate')
        
    def test_datasrouce_manager_next_item(self):
        ds = DatasourceManager('test.xlsx', 'Source','HRMI01')
        a = ds.get_var('LS_2ADHESION', 'CODE_FORMATION')
        self.assertEqual(a.value.value, 'HRMI0101')
        
        ds.next_item()
        a = ds.get_var('LS_2ADHESION', 'CODE_FORMATION')
        self.assertEqual(a.value.value, 'HRMI0102')
        
class FormulaGetVarTest(unittest.TestCase):
    def test_formula_getvar(self):
        a = Formula('VarTest','(Fiche_Client.Dépla+(Fiche_Client.totalrestauHT*Fiche_Client.coutformHT))/Fiche_Client.Txtva','Monétaire').get_var_names()
        self.assertEqual(a,['Fiche_Client.Dépla', 'Fiche_Client.totalrestauHT', 'Fiche_Client.coutformHT', 'Fiche_Client.Txtva'])
        
        b = Formula('Var201','Fiche_Client.DateCDC','Date').get_var_names()
        self.assertEqual(b,['Fiche_Client.DateCDC'])
        
        c = Formula('Var201','Var203','Date').get_var_names()
        self.assertEqual(c, ['Var203'])
        
class FormulaManagerTest(unittest.TestCase) :
    def test_read_spec(self):
        datasource = DatasourceManager('test.xlsx', 'Source','HRMI01')
        formula_manager = FormulaManager('test.xlsx', 'Calcul', datasource)
        self.assertEqual(len(formula_manager.formulas),93)
        
    def test_calcul(self):
        datasource = DatasourceManager('test_spec.xlsx', 'Source','HRMI01')
        formula_manager = FormulaManager('test_spec.xlsx', 'Calcul', datasource)
        dic = formula_manager.next_vars_dictionary()
        self.assertEqual(dic['Var1'],'31290')
        self.assertEqual(dic['Var2'],'13/01/2021')        
        self.assertEqual(dic['Var3'],'titi')
        self.assertEqual(dic['Var4'],'31 280')
        
    def test_calcul_2(self):
        pass
    
    def test_next_vars_dict_formulamanager(self):
        ds = DatasourceManager('test.xlsx', 'Source','HRMI01')
        fm = FormulaManager('test.xlsx','Calcul',ds)
        d = fm.next_vars_dictionary()
        self.assertEqual(d['Var652'],'1')
        self.assertEqual(d['Var602'],'HRMI0101')
        
        e = fm.next_vars_dictionary()
        self.assertEqual(e['Var652'],'1')
        self.assertEqual(e['Var602'],'HRMI0102')
        
class Composer(unittest.TestCase):
    pass

if __name__ == '__main__' :
    unittest.main()


 
