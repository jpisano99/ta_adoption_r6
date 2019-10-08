class Customer:
    def __init__(self, cust_id):
        self.cust_id = cust_id
        self.pss = ''
        self.tsa = ''
        self.am = ''
        self.sales_lev_1 = ''
        self.sales_lev_2 = ''
        self.sales_lev_3 = ''
        self.sales_lev_4 = ''
        self.sales_lev_5 = ''
        self.sales_lev_6 = ''

        self.aliases = []  # Simple list of customer erp names for this cust id
        self.orders = {}  # Simple dict of {SO #: [sku1, sku2]}
        self.as_pids = {}  # Simple dict of {SO #: [(pid1, as_customer_name)]}
        self.subs = []  # Simple list of lists of subscriptions[]
        self.saas = {}

    def add_order(self, order_num, sku):
        if order_num in self.orders:
            sku_list = self.orders[order_num]
            sku_list.append(sku)
            self.orders[order_num] = sku_list
        else:
            self.orders[order_num] = [sku]

    def add_alias(self, erp_name):
        add_it = True
        for alias in self.aliases:
            if alias == erp_name:
                add_it = False
                break
        if add_it is True:
            self.aliases.append(erp_name)

    def add_as_pid(self, order_num, as_pid_info):
        # as_pid_info is a full list of tuples (as_pid, as_cust_name)
        # with ALL PIDs for this SO for this cust_id
        self.as_pids[order_num] = as_pid_info

    def add_sub_id(self, sub_info):
        # sub_id_pid_info is a list of lists (sub_id, sub_cust_name, start_date, renew_date, status, monthly_rev)
        # with ALL Subscriptions for this erp_customer name
        self.subs.append(sub_info)

