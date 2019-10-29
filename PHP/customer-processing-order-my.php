<?php
/**
 * Customer processing order email
 *
 * This template can be overridden by copying it to yourtheme/woocommerce/emails/customer-processing-order.php.
 *
 * HOWEVER, on occasion WooCommerce will need to update template files and you
 * (the theme developer) will need to copy the new files to your theme to
 * maintain compatibility. We try to do this as little as possible, but it does
 * happen. When this occurs the version of the template file will be bumped and
 * the readme will list any important changes.
 *
 * @see https://docs.woocommerce.com/document/template-structure/
 * @package WooCommerce/Templates/Emails
 * @version 3.5.4
 */

if ( ! defined( 'ABSPATH' ) ) {
	exit;
}

/*
 * @hooked WC_Emails::email_header() Output the email header
 */
do_action( 'woocommerce_email_header', $email_heading, $email ); ?>

<?php /* translators: %s: Customer first name */ ?>
<p><?php printf( esc_html__( 'Здравствуйте, %s.', 'woocommerce' ), esc_html( $order->get_billing_first_name() ) ); ?></p>
<?php /* translators: %s: Order number */ ?>
<p><?php printf( esc_html__('Мы получили ваш заказ &mdash; %s. Способ оплаты: %s.', 'woocommerce' ), esc_html( $order->get_order_number() ),esc_html( $order->get_payment_method_title() ) ); ?></p>
<p><?php printf( 'Наши менеджеры свяжутся с вами по указанному номеру, подтвердят заказ и согласуют доставку.')?></p>
<p><?php printf( 'Если по какой-то причине этого не произойдет, либо вы не можете в данный момент подтвердить заказ по телефону, свяжитесь, пожалуйста, с нами любым удобным способом:')?></p>
<p>По телефону: <br>
<a href="tel:+7-499-490-66-81">+7-499-490-66-81</a> (Москва)<br>
<a href="tel:+7-800-350-90-48">+7-800-350-90-48</a> (Остальная Россия)<br><br>
По почте: <a href="mailto:mail@cacaoclub.ru">mail@cacaoclub.ru</a><br>				    
Instagram: <a href="https://www.instagram.com/cacaoclubshop" target="_blank">@cacaoclubshop</a><br>
VK: <a href="https://vk.com/cacaoclubshop" target="_blank">@cacaoclubshop</a></p>    

<?php

/*
 * @hooked WC_Emails::order_details() Shows the order details table.
 * @hooked WC_Structured_Data::generate_order_data() Generates structured data.
 * @hooked WC_Structured_Data::output_structured_data() Outputs structured data.
 * @since 2.5.0
 */
do_action( 'woocommerce_email_order_details_custom', $order, $sent_to_admin, $plain_text, $email );

/*
 * @hooked WC_Emails::order_meta() Shows order meta data.
 */
do_action( 'woocommerce_email_order_meta', $order, $sent_to_admin, $plain_text, $email );

/*
 * @hooked WC_Emails::customer_details() Shows customer details
 * @hooked WC_Emails::email_address() Shows email address
 */
do_action( 'woocommerce_email_customer_details_custom', $order, $sent_to_admin, $plain_text, $email );

?>
<p>
    Подписывайтесь на наши соцсети <a href="https://www.instagram.com/cacaoclubshop/" target="_blank" rel="noopener noreferrer">
			<img alt="instagram" title="instagram" src="http://test.cacaoclub.ru/wp-content/uploads/2018/10/instagram.png"></a>  и   
<a href="https://vk.com/cacaoclubshop" target="_blank" rel="noopener noreferrer"><img alt="vkontakte" title="vkontakte" src="http://test.cacaoclub.ru/wp-content/uploads/2018/10/vk.png"></a>, там регулярно появляются новые рецепты и полезные статьи!
</p>

<?php
/*
 * @hooked WC_Emails::email_footer() Output the email footer
 */
 do_action( 'woocommerce_email_footer_custom', $email );
?>